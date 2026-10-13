// Ask who is presenting and write presenter.json, which the cover and the
// closing slide read. Run again to change it; Enter keeps the current value.
//   npm run setup
const fs = require("fs");
const readline = require("readline");
const { PRESENTER_FIELDS, PRESENTER_FILE } = require("../template");

const LABELS = {
  company: "회사",
  department: "부서",
  name: "이름",
  position: "직책",
  email: "이메일 (선택)",
  phone: "연락처 (선택)",
};

async function main() {
  let current = {};
  if (fs.existsSync(PRESENTER_FILE)) current = JSON.parse(fs.readFileSync(PRESENTER_FILE, "utf8"));

  // Read lines through the iterator rather than rl.question(), which drops
  // lines that arrive before it is asked (piped input).
  const rl = readline.createInterface({ input: process.stdin });
  const lines = rl[Symbol.asyncIterator]();
  const next = {};
  console.log("발표자 정보 — Enter 는 현재 값 유지, '-' 는 비우기\n");
  for (const key of PRESENTER_FIELDS) {
    const now = current[key] || "";
    process.stdout.write(`${LABELS[key]}${now ? ` [${now}]` : ""}: `);
    const { value, done } = await lines.next();
    if (!process.stdin.isTTY) process.stdout.write("\n");
    const answer = done ? "" : value.trim();
    next[key] = answer === "-" ? "" : answer || now;
  }
  rl.close();

  fs.writeFileSync(PRESENTER_FILE, JSON.stringify(next, null, 2) + "\n");
  console.log(`\n${PRESENTER_FILE} 에 저장했습니다 (git 에는 올라가지 않습니다).`);
}

main().catch((err) => {
  console.error(err.message);
  process.exit(1);
});
