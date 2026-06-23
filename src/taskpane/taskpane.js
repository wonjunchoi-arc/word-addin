/* global document, fetch, Blob, FormData, Office, Word */

const API = "https://localhost:8000";

// /document-outline 응답을 보관해 섹션 선택 시 locator를 조회한다.
let outlineFlatItems = [];

Office.onReady(() => {
  document.getElementById("run-btn").onclick = runPipeline;
  document.getElementById("outline-btn").onclick = loadOutline;
  document.getElementById("outline-select").onchange = onOutlineSelect;
  loadCases();
});

async function loadCases() {
  setStatus("사건 목록 불러오는 중...");
  try {
    const res = await fetch(`${API}/cases`);
    if (!res.ok) throw new Error(`서버 오류: ${res.status}`);
    const cases = await res.json();
    renderCases(cases);
    setStatus(`${cases.length}건 로드됨`);
  } catch (e) {
    setStatus(`오류: ${e.message}`);
  }
}

function renderCases(cases) {
  const list = document.getElementById("case-list");
  const tpl = document.getElementById("case-item-tpl");
  list.innerHTML = "";

  cases.forEach((c) => {
    const item = tpl.content.cloneNode(true);
    item.querySelector(".case-title").textContent = `${c.subject_id} — ${c.event_pt}`;
    item.querySelector(".case-meta").textContent = `${c.arm} | ${c.soc}`;
    const btn = item.querySelector(".insert-btn");
    btn.onclick = () => insertCase(c.subject_id, btn);
    list.appendChild(item);
  });
}

async function insertCase(subjectId, btn) {
  btn.disabled = true;
  btn.textContent = "삽입 중...";
  setStatus(`${subjectId} 삽입 중...`);

  try {
    const res = await fetch(`${API}/cases/${subjectId}`);
    if (!res.ok) throw new Error(`서버 오류: ${res.status}`);
    const data = await res.json();

    await Word.run(async (context) => {
      context.document.body.insertHtml(data.html, Word.InsertLocation.end);
      await context.sync();
    });

    setStatus(`${subjectId} 삽입 완료`);
    btn.textContent = "삽입 완료 ✓";
  } catch (e) {
    setStatus(`오류: ${e.message}`);
    btn.disabled = false;
    btn.textContent = "문서에 삽입";
  }
}

async function runPipeline() {
  const btn = document.getElementById("run-btn");
  btn.disabled = true;
  btn.textContent = "실행 중...";
  setStatus("파이프라인 실행 중...");

  try {
    const res = await fetch(`${API}/run`, { method: "POST" });
    if (!res.ok) throw new Error(`서버 오류: ${res.status}`);
    setStatus("완료. 사건 목록 새로고침 중...");
    await loadCases();
  } catch (e) {
    setStatus(`오류: ${e.message}`);
  } finally {
    btn.disabled = false;
    btn.textContent = "전체 파이프라인 실행";
  }
}

// ---- 문서 구조(outline) ----

// "현재 문서 구조 분석" 버튼: 열린 문서를 백엔드로 보내 outline을 받아 드롭다운을 채운다.
async function loadOutline() {
  const btn = document.getElementById("outline-btn");
  btn.disabled = true;
  btn.textContent = "분석 중...";
  setStatus("현재 문서를 읽는 중...");

  try {
    const docBytes = await getDocumentBytes();

    setStatus("문서 구조 분석 중...");
    const form = new FormData();
    form.append("file", new Blob([docBytes]), "document.docx");

    const res = await fetch(`${API}/document-outline`, { method: "POST", body: form });
    if (!res.ok) throw new Error(`서버 오류: ${res.status}`);
    const data = await res.json();

    outlineFlatItems = data.flat_items || [];
    renderOutline(outlineFlatItems);
    setStatus(`${outlineFlatItems.length}개 섹션 로드됨`);
  } catch (e) {
    setStatus(`오류: ${e.message}`);
  } finally {
    btn.disabled = false;
    btn.textContent = "현재 문서 구조 분석";
  }
}

// Office.js getFileAsync로 열린 문서를 docx 바이너리(Uint8Array)로 모은다.
// 문서는 슬라이스로 쪼개져 오므로 순서대로 읽어 이어 붙인다.
function getDocumentBytes() {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 65536 },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error.message));
          return;
        }

        const file = result.value;
        const sliceCount = file.sliceCount;
        const slices = new Array(sliceCount);
        let received = 0;

        const readSlice = (index) => {
          file.getSliceAsync(index, (sliceResult) => {
            if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync(() => {});
              reject(new Error(sliceResult.error.message));
              return;
            }

            slices[index] = sliceResult.value.data; // number[] (byte 값)
            received += 1;

            if (received === sliceCount) {
              file.closeAsync(() => {});
              resolve(mergeSlices(slices));
            } else if (index + 1 < sliceCount) {
              readSlice(index + 1);
            }
          });
        };

        if (sliceCount === 0) {
          file.closeAsync(() => {});
          reject(new Error("문서가 비어 있습니다."));
          return;
        }
        readSlice(0);
      }
    );
  });
}

function mergeSlices(slices) {
  const total = slices.reduce((sum, s) => sum + s.length, 0);
  const bytes = new Uint8Array(total);
  let offset = 0;
  slices.forEach((s) => {
    bytes.set(s, offset);
    offset += s.length;
  });
  return bytes;
}

function renderOutline(items) {
  const select = document.getElementById("outline-select");
  select.innerHTML = '<option value="">— 섹션 선택 —</option>';

  items.forEach((item) => {
    const opt = document.createElement("option");
    opt.value = item.section_id;
    // level 깊이만큼 들여쓰기해 계층을 표현한다.
    const indent = "  ".repeat(Math.max(0, (item.level || 1) - 1));
    opt.textContent = indent + (item.display_text || item.title || item.section_id);
    select.appendChild(opt);
  });

  document.getElementById("outline-picker").hidden = items.length === 0;
}

// 드롭다운 선택 시 해당 섹션의 heading으로 Word 커서를 이동한다.
async function onOutlineSelect(e) {
  const sectionId = e.target.value;
  if (!sectionId) return;

  const item = outlineFlatItems.find((i) => i.section_id === sectionId);
  if (!item || !item.heading_text) {
    setStatus("이동할 섹션 정보를 찾을 수 없습니다.");
    return;
  }

  setStatus(`'${item.display_text || item.title}'(으)로 이동 중...`);

  try {
    await Word.run(async (context) => {
      const results = context.document.body.search(item.heading_text, { matchCase: false });
      results.load("items");
      await context.sync();

      if (results.items.length === 0) {
        setStatus("본문에서 해당 섹션을 찾지 못했습니다.");
        return;
      }

      // 검색 결과는 같은 제목이 목차(TOC)와 실제 본문 양쪽에 잡힌다.
      // 각 결과가 속한 단락의 스타일을 확인해 실제 heading 문단만 고른다.
      const paragraphs = results.items.map((r) => r.paragraphs.getFirst());
      paragraphs.forEach((p) => p.load("styleBuiltIn,style"));
      await context.sync();

      // 목차 항목은 'Toc*' 스타일, 실제 제목은 'Heading*' 스타일.
      const isHeading = (p) => {
        const built = (p.styleBuiltIn || "").toLowerCase();
        const name = (p.style || "").toLowerCase();
        const heading =
          built.startsWith("heading") || name.includes("heading") || name.includes("제목");
        const toc = built.startsWith("toc") || name.includes("toc") || name.includes("목차");
        return heading && !toc;
      };

      // 1순위: heading 스타일인 결과. 없으면 목차가 아닌 마지막 결과(보통 본문)로 fallback.
      let targetIndex = paragraphs.findIndex(isHeading);
      if (targetIndex === -1) {
        const nonToc = paragraphs
          .map((p, i) => ({ p, i }))
          .filter(({ p }) => {
            const built = (p.styleBuiltIn || "").toLowerCase();
            const name = (p.style || "").toLowerCase();
            return !(built.startsWith("toc") || name.includes("toc") || name.includes("목차"));
          });
        targetIndex = nonToc.length ? nonToc[nonToc.length - 1].i : results.items.length - 1;
      }

      const target = results.items[targetIndex];
      target.select(Word.SelectionMode.start);
      await context.sync();
      setStatus(`'${item.display_text || item.title}'(으)로 이동 완료`);
    });
  } catch (err) {
    setStatus(`오류: ${err.message}`);
  }
}

function setStatus(msg) {
  document.getElementById("status").textContent = msg;
}
