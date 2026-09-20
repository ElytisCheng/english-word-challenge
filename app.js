"use strict";

const PROGRESS_STORAGE_KEY = "englishWordChallenge.progress.v1";
const DAILY_STORAGE_KEY = "englishWordChallenge.daily.v2";
const REVIEW_INTERVAL_DAYS = 30;

const elements = {
  setupPanel: document.querySelector("#setupPanel"),
  gamePanel: document.querySelector("#gamePanel"),
  resultPanel: document.querySelector("#resultPanel"),
  wordFile: document.querySelector("#wordFile"),
  fileLabel: document.querySelector("#fileLabel"),
  loadStatus: document.querySelector("#loadStatus"),
  questionCount: document.querySelector("#questionCount"),
  startButton: document.querySelector("#startButton"),
  progressText: document.querySelector("#progressText"),
  scoreText: document.querySelector("#scoreText"),
  mixText: document.querySelector("#mixText"),
  progressBar: document.querySelector("#progressBar"),
  chinesePrompt: document.querySelector("#chinesePrompt"),
  hintOutput: document.querySelector("#hintOutput"),
  answerForm: document.querySelector("#answerForm"),
  answerInput: document.querySelector("#answerInput"),
  feedback: document.querySelector("#feedback"),
  hintButton: document.querySelector("#hintButton"),
  speakButton: document.querySelector("#speakButton"),
  skipButton: document.querySelector("#skipButton"),
  quitButton: document.querySelector("#quitButton"),
  resultSummary: document.querySelector("#resultSummary"),
  accuracyScore: document.querySelector("#accuracyScore"),
  scoreGrade: document.querySelector("#scoreGrade"),
  correctResultCount: document.querySelector("#correctResultCount"),
  wrongResultCount: document.querySelector("#wrongResultCount"),
  correctList: document.querySelector("#correctList"),
  mistakeList: document.querySelector("#mistakeList"),
  retryButton: document.querySelector("#retryButton"),
  changeFileButton: document.querySelector("#changeFileButton")
};

const state = {
  vocabulary: [],
  questions: [],
  currentIndex: 0,
  correctCount: 0,
  attempts: 0,
  mistakes: [],
  roundResults: [],
  acceptingAnswer: true,
  lastSpokenAnswer: "",
  currentHadMistake: false,
  currentAttempts: 0
};

elements.wordFile.addEventListener("change", handleFileSelection);
elements.startButton.addEventListener("click", startGame);
elements.answerForm.addEventListener("submit", checkAnswer);
elements.hintButton.addEventListener("click", showHint);
elements.speakButton.addEventListener("click", () => speakEnglish(state.lastSpokenAnswer));
elements.skipButton.addEventListener("click", skipQuestion);
elements.quitButton.addEventListener("click", finishGame);
elements.retryButton.addEventListener("click", startGame);
elements.changeFileButton.addEventListener("click", resetToSetup);

loadDefaultVocabulary();

async function handleFileSelection(event) {
  const file = event.target.files[0];
  elements.startButton.disabled = true;
  elements.loadStatus.className = "status";

  if (!file) {
    elements.fileLabel.textContent = "选择 Word 或 Excel 词库";
    elements.loadStatus.textContent = "";
    return;
  }

  elements.fileLabel.textContent = file.name;
  const extension = file.name.toLowerCase().split(".").pop();

  if (!["docx", "xlsx", "xlsm", "xls"].includes(extension)) {
    showLoadError("请选择 .docx、.xlsx、.xlsm 或 .xls 文件。");
    return;
  }

  try {
    elements.loadStatus.textContent = "正在读取词库...";
    const entries = extension === "docx"
      ? await extractWordVocabulary(file)
      : await extractExcelVocabulary(file);

    if (entries.length === 0) {
      showLoadError("没有识别到词条。请使用英文、中文两列；Word 也支持“英文 = 中文”格式。");
      return;
    }

    applyVocabulary(entries, file.name, false);
  } catch (error) {
    console.error(error);
    if (/Encrypted file|EncryptionInfo|DRM|password/i.test(error.message)) {
      showLoadError("该文件受 Microsoft 组织 DRM/敏感度标签保护，浏览器无法直接解密。请先用 Excel 另存为标准 .xlsx，或使用程序目录中的转换工具。");
    } else {
      showLoadError("词库读取失败。请确认文件未损坏、未受密码保护，并使用受支持的格式。");
    }
  }
}

function loadDefaultVocabulary() {
  const defaultVocabulary = window.DEFAULT_VOCABULARY;
  if (!defaultVocabulary?.entries || !Array.isArray(defaultVocabulary.entries)) {
    elements.loadStatus.className = "status error";
    elements.loadStatus.textContent =
      "未生成默认词库数据。请使用“启动英文词汇闯关.cmd”启动，或手动选择词库。";
    return;
  }

  const entries = deduplicateEntries(defaultVocabulary.entries
    .map(entry => createEntry(cleanText(entry.english), cleanText(entry.chinese)))
    .filter(entry => entry.english && entry.chinese));

  if (entries.length === 0) {
    showLoadError("默认 Excel 中没有识别到有效的英文和中文词条。");
    return;
  }

  applyVocabulary(entries, defaultVocabulary.source || "单词&词组.xlsx", true);
}

function applyVocabulary(entries, sourceName, automatic) {
  state.vocabulary = entries;
  elements.questionCount.max = String(Math.max(entries.length, 1));
  elements.questionCount.value = String(Math.min(50, entries.length));
  const poolSummary = getPoolSummary(entries);
  elements.fileLabel.textContent = automatic
    ? `已自动读取：${sourceName}`
    : sourceName;
  elements.loadStatus.className = "status success";
  elements.loadStatus.textContent =
    `已读取 ${entries.length} 个词条：${poolSummary.unseen} 个生词，${poolSummary.mistakes} 个错题，${poolSummary.due} 个到期复习词。`;
  elements.startButton.disabled = false;
}

async function extractWordVocabulary(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const documentPart = zip.file("word/document.xml");

  if (!documentPart) {
    throw new Error("word/document.xml was not found");
  }

  const xmlText = await documentPart.async("string");
  const xml = new DOMParser().parseFromString(xmlText, "application/xml");

  if (xml.querySelector("parsererror")) {
    throw new Error("Invalid Word XML");
  }

  const entries = [
    ...extractTableEntries(xml),
    ...extractParagraphEntries(xml)
  ];

  return deduplicateEntries(entries);
}

async function extractExcelVocabulary(file) {
  const workbook = XLSX.read(await file.arrayBuffer(), {
    type: "array",
    cellText: true,
    cellDates: false
  });
  const entries = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: false,
      defval: ""
    });
    for (const row of rows) {
      const values = row
        .map(value => String(value))
        .map(cleanText)
        .filter(Boolean);
      const pair = findEnglishChinesePair(values);

      if (pair && !looksLikeHeader(pair.english, pair.chinese)) {
        entries.push(pair);
      }
    }
  }

  return deduplicateEntries(entries);
}

function findEnglishChinesePair(values) {
  for (let firstIndex = 0; firstIndex < values.length; firstIndex += 1) {
    for (let secondIndex = firstIndex + 1; secondIndex < values.length; secondIndex += 1) {
      const pair = identifyEnglishChinese(values[firstIndex], values[secondIndex]);
      if (pair) {
        return pair;
      }
    }
  }
  return null;
}

function extractTableEntries(xml) {
  const entries = [];
  const rows = Array.from(xml.getElementsByTagNameNS("*", "tr"));

  for (const row of rows) {
    const cells = Array.from(row.children)
      .filter(node => node.localName === "tc")
      .map(readWordNodeText)
      .map(cleanText)
      .filter(Boolean);

    if (cells.length < 2) {
      continue;
    }

    const pair = identifyEnglishChinese(cells[0], cells[1]);
    if (pair && !looksLikeHeader(pair.english, pair.chinese)) {
      entries.push(pair);
    }
  }

  return entries;
}

function extractParagraphEntries(xml) {
  const entries = [];
  const paragraphs = Array.from(xml.getElementsByTagNameNS("*", "p"));

  for (const paragraph of paragraphs) {
    if (hasAncestor(paragraph, "tbl")) {
      continue;
    }

    const line = cleanText(readWordNodeText(paragraph));
    if (!line) {
      continue;
    }

    const parts = line.split(/\s*(?:\t|=|＝|\||：|:|—|–)\s*/).filter(Boolean);
    if (parts.length < 2) {
      continue;
    }

    const pair = identifyEnglishChinese(parts[0], parts.slice(1).join(" "));
    if (pair && !looksLikeHeader(pair.english, pair.chinese)) {
      entries.push(pair);
    }
  }

  return entries;
}

function readWordNodeText(root) {
  const pieces = [];

  function visit(node) {
    if (node.nodeType === Node.ELEMENT_NODE) {
      if (node.localName === "t") {
        pieces.push(node.textContent || "");
        return;
      }
      if (node.localName === "tab") {
        pieces.push("\t");
        return;
      }
      if (node.localName === "br" || node.localName === "cr") {
        pieces.push(" ");
        return;
      }
    }

    for (const child of node.childNodes) {
      visit(child);
    }
  }

  visit(root);
  return pieces.join("");
}

function hasAncestor(node, localName) {
  let current = node.parentElement;
  while (current) {
    if (current.localName === localName) {
      return true;
    }
    current = current.parentElement;
  }
  return false;
}

function identifyEnglishChinese(first, second) {
  const firstHasEnglish = /[A-Za-z]/.test(first) && !looksLikePronunciation(first);
  const secondHasEnglish = /[A-Za-z]/.test(second) && !looksLikePronunciation(second);
  const firstHasChinese = /[\u3400-\u9fff]/.test(first);
  const secondHasChinese = /[\u3400-\u9fff]/.test(second);

  if (firstHasEnglish && secondHasChinese) {
    return createEntry(first, second);
  }
  if (secondHasEnglish && firstHasChinese) {
    return createEntry(second, first);
  }
  return null;
}

function looksLikePronunciation(value) {
  return /^\s*(?:\/.*\/|\[.*\])\s*$/.test(value);
}

function createEntry(english, chinese) {
  const answers = expandAnswerVariants(english);

  return {
    english: answers[0],
    answers,
    chinese: cleanText(chinese)
  };
}

function expandAnswerVariants(value) {
  const variants = [];
  const alternatives = value
    .split(/\s*[;；]\s*/)
    .map(cleanText)
    .filter(Boolean);

  for (const alternative of alternatives) {
    const phraseVariants = expandSpecialSlashPatterns(alternative);
    for (const phrase of phraseVariants) {
      const spacedAlternatives = /\s+\/\s+/.test(phrase)
        ? phrase.split(/\s+\/\s+/).map(cleanText).filter(Boolean)
        : [phrase];
      for (const spacedAlternative of spacedAlternatives) {
        variants.push(...expandTokenSlashes(spacedAlternative));
      }
    }
  }

  const unique = new Map();
  for (const variant of variants) {
    const cleaned = cleanText(variant);
    const key = normalizeAnswer(cleaned);
    if (key && !unique.has(key)) {
      unique.set(key, cleaned);
    }
  }
  return [...unique.values()];
}

function expandSpecialSlashPatterns(value) {
  const patterns = [
    {
      expression: /doing\s+sth\.\s*\/\s*to\s+do\s+sth\./i,
      replacements: ["doing sth.", "to do sth."]
    },
    {
      expression: /do\s+sth\.\s*\/\s*doing\s+sth\./i,
      replacements: ["do sth.", "doing sth."]
    }
  ];

  let variants = [value];
  for (const pattern of patterns) {
    const expanded = [];
    for (const variant of variants) {
      if (pattern.expression.test(variant)) {
        for (const replacement of pattern.replacements) {
          expanded.push(variant.replace(pattern.expression, replacement));
        }
      } else {
        expanded.push(variant);
      }
    }
    variants = expanded;
  }
  return variants;
}

function expandTokenSlashes(value) {
  let variants = [value];
  const tokenSlash = /([A-Za-z][A-Za-z'.-]*\.?)\s*\/\s*([A-Za-z][A-Za-z'.-]*\.?)/;

  for (let pass = 0; pass < 8; pass += 1) {
    let expandedAny = false;
    const expanded = [];

    for (const variant of variants) {
      const match = variant.match(tokenSlash);
      if (!match) {
        expanded.push(variant);
        continue;
      }

      expandedAny = true;
      const before = variant.slice(0, match.index);
      const after = variant.slice(match.index + match[0].length);
      expanded.push(`${before}${match[1]}${after}`);
      expanded.push(`${before}${match[2]}${after}`);
    }

    variants = expanded;
    if (!expandedAny) {
      break;
    }
  }

  return variants;
}

function looksLikeHeader(english, chinese) {
  return /^(english|word|phrase|英文|单词|词组)$/i.test(english)
    || /^(chinese|meaning|中文|释义|含义)$/i.test(chinese);
}

function deduplicateEntries(entries) {
  const seen = new Set();
  return entries.filter(entry => {
    if (!entry.english || !entry.chinese) {
      return false;
    }
    const key = `${normalizeAnswer(entry.english)}|${entry.chinese}`;
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function cleanText(value) {
  return value.replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
}

function normalizeAnswer(value) {
  return value
    .normalize("NFKC")
    .toLowerCase()
    .replace(/[’‘]/g, "'")
    .replace(/[“”]/g, "\"")
    .replace(/\s+/g, " ")
    .replace(/^[\s.,!?;:"']+|[\s.,!?;:"']+$/g, "")
    .trim();
}

function startGame() {
  if (state.vocabulary.length === 0) {
    return;
  }

  const requestedCount = Number.parseInt(elements.questionCount.value, 10) || 50;
  const count = Math.min(Math.max(requestedCount, 1), state.vocabulary.length);

  state.questions = selectDailyQuestions(state.vocabulary, count);
  state.currentIndex = 0;
  state.correctCount = 0;
  state.attempts = 0;
  state.mistakes = [];
  state.roundResults = [];
  state.acceptingAnswer = true;
  state.lastSpokenAnswer = "";
  state.currentHadMistake = false;
  state.currentAttempts = 0;

  showPanel(elements.gamePanel);
  renderQuestion();
}

function renderQuestion() {
  if (state.currentIndex >= state.questions.length) {
    finishGame();
    return;
  }

  const question = state.questions[state.currentIndex];
  const displayedNumber = state.currentIndex + 1;
  const mix = countQuestionTypes(state.questions);

  elements.progressText.textContent = `第 ${displayedNumber} / ${state.questions.length} 题`;
  elements.scoreText.textContent = `答对 ${state.correctCount} 题`;
  elements.mixText.textContent =
    `今日生词 ${mix.new} · 错题 ${mix.mistake} · 到期复习 ${mix.due} · 普通复习 ${mix.review}`;
  elements.progressBar.style.width = `${(state.currentIndex / state.questions.length) * 100}%`;
  elements.chinesePrompt.textContent = question.chinese;
  elements.hintOutput.textContent = "";
  elements.feedback.textContent = "";
  elements.feedback.className = "feedback";
  elements.answerInput.value = "";
  elements.answerInput.disabled = false;
  elements.hintButton.disabled = false;
  elements.skipButton.disabled = false;
  elements.speakButton.disabled = !state.lastSpokenAnswer;
  state.acceptingAnswer = true;
  state.currentHadMistake = false;
  state.currentAttempts = 0;
  recordSeen(question);

  requestAnimationFrame(() => elements.answerInput.focus());
}

async function checkAnswer(event) {
  event.preventDefault();

  if (!state.acceptingAnswer) {
    return;
  }

  const userAnswer = normalizeAnswer(elements.answerInput.value);
  if (!userAnswer) {
    elements.feedback.textContent = "请先输入英文答案。";
    elements.feedback.className = "feedback wrong";
    return;
  }

  const question = state.questions[state.currentIndex];
  const isCorrect = question.answers.some(answer => normalizeAnswer(answer) === userAnswer);
  state.attempts += 1;
  state.currentAttempts += 1;

  if (!isCorrect) {
    if (!state.currentHadMistake) {
      state.currentHadMistake = true;
      state.mistakes.push(question);
      recordMistake(question);
    }
    elements.feedback.textContent = "还不正确，再试一次。";
    elements.feedback.className = "feedback wrong";
    elements.answerInput.select();
    return;
  }

  const firstTryCorrect = !state.currentHadMistake;
  if (firstTryCorrect) {
    state.correctCount += 1;
  }
  recordCorrect(question, firstTryCorrect);
  state.roundResults.push({
    question,
    outcome: firstTryCorrect ? "correct" : "wrong",
    attempts: state.currentAttempts
  });
  state.acceptingAnswer = false;
  state.lastSpokenAnswer = question.english;
  elements.feedback.textContent = `正确！${question.english}`;
  elements.feedback.className = "feedback correct";
  elements.answerInput.disabled = true;
  elements.hintButton.disabled = true;
  elements.skipButton.disabled = true;
  elements.speakButton.disabled = true;
  playCorrectSound();

  const answeredIndex = state.currentIndex;
  await wait(300);
  await speakEnglish(question.english);
  elements.feedback.textContent = `正确！${question.english}　即将进入下一题…`;
  await wait(1200);

  if (state.currentIndex === answeredIndex && !elements.gamePanel.classList.contains("hidden")) {
    state.currentIndex += 1;
    renderQuestion();
  }
}

function showHint() {
  const question = state.questions[state.currentIndex];
  if (!question) {
    return;
  }

  elements.hintOutput.textContent = question.english
    .split(/\s+/)
    .map(word => `${word.charAt(0)}${"_".repeat(Math.max(word.length - 1, 0))}`)
    .join(" ");
  elements.answerInput.focus();
}

async function skipQuestion() {
  const question = state.questions[state.currentIndex];
  if (!question || !state.acceptingAnswer) {
    return;
  }

  if (!state.currentHadMistake) {
    state.currentHadMistake = true;
    state.mistakes.push(question);
    recordMistake(question);
  }
  state.roundResults.push({
    question,
    outcome: "wrong",
    attempts: state.currentAttempts,
    skipped: true
  });
  state.acceptingAnswer = false;
  state.lastSpokenAnswer = question.english;
  elements.feedback.textContent = `答案：${question.english}`;
  elements.feedback.className = "feedback wrong";
  elements.answerInput.disabled = true;
  elements.hintButton.disabled = true;
  elements.skipButton.disabled = true;
  elements.speakButton.disabled = true;

  const answeredIndex = state.currentIndex;
  await speakEnglish(question.english);
  elements.feedback.textContent = `答案：${question.english}　即将进入下一题…`;
  await wait(1200);

  if (state.currentIndex === answeredIndex && !elements.gamePanel.classList.contains("hidden")) {
    state.currentIndex += 1;
    renderQuestion();
  }
}

function speakEnglish(text) {
  return new Promise(resolve => {
    if (!text || !("speechSynthesis" in window)) {
      resolve();
      return;
    }

    window.speechSynthesis.cancel();
    const utterance = new SpeechSynthesisUtterance(text);
    const voices = window.speechSynthesis.getVoices();
    utterance.voice = voices.find(voice => /^en-US$/i.test(voice.lang))
      || voices.find(voice => /^en/i.test(voice.lang))
      || null;
    utterance.lang = utterance.voice?.lang || "en-US";
    utterance.rate = 0.82;
    utterance.pitch = 1;

    let completed = false;
    let fallbackTimer;
    const finish = () => {
      if (completed) {
        return;
      }
      completed = true;
      window.clearTimeout(fallbackTimer);
      resolve();
    };
    fallbackTimer = window.setTimeout(finish, Math.max(5000, text.length * 260));

    utterance.onend = finish;
    utterance.onerror = finish;
    window.speechSynthesis.speak(utterance);
  });
}

function playCorrectSound() {
  const AudioContextClass = window.AudioContext || window.webkitAudioContext;
  if (!AudioContextClass) {
    return;
  }

  const context = new AudioContextClass();
  const masterGain = context.createGain();
  masterGain.gain.setValueAtTime(0.0001, context.currentTime);
  masterGain.gain.exponentialRampToValueAtTime(0.16, context.currentTime + 0.02);
  masterGain.gain.exponentialRampToValueAtTime(0.0001, context.currentTime + 0.55);
  masterGain.connect(context.destination);

  const notes = [
    { frequency: 659.25, start: 0, duration: 0.18 },
    { frequency: 783.99, start: 0.14, duration: 0.2 },
    { frequency: 1046.5, start: 0.3, duration: 0.24 }
  ];

  for (const note of notes) {
    const oscillator = context.createOscillator();
    const noteGain = context.createGain();
    oscillator.type = "sine";
    oscillator.frequency.value = note.frequency;
    noteGain.gain.setValueAtTime(0.7, context.currentTime + note.start);
    noteGain.gain.exponentialRampToValueAtTime(
      0.0001,
      context.currentTime + note.start + note.duration
    );
    oscillator.connect(noteGain);
    noteGain.connect(masterGain);
    oscillator.start(context.currentTime + note.start);
    oscillator.stop(context.currentTime + note.start + note.duration);
  }

  window.setTimeout(() => context.close(), 800);
}

function wait(milliseconds) {
  return new Promise(resolve => window.setTimeout(resolve, milliseconds));
}

function selectDailyQuestions(vocabulary, count) {
  const vocabularyByKey = new Map(vocabulary.map(entry => [entryKey(entry), entry]));
  const storedDaily = loadJson(DAILY_STORAGE_KEY, null);
  const today = getLocalDateKey();

  if (storedDaily?.date === today
      && storedDaily.count === count
      && Array.isArray(storedDaily.items)
      && storedDaily.items.length === count
      && storedDaily.items.every(item => vocabularyByKey.has(item.key))) {
    return storedDaily.items.map(item => ({
      ...vocabularyByKey.get(item.key),
      challengeType: item.type
    }));
  }

  const progress = loadProgress();
  const unseen = [];
  const mistakes = [];
  const due = [];
  const review = [];

  for (const entry of vocabulary) {
    const stats = progress[entryKey(entry)];
    if (!stats || stats.seenCount === 0) {
      unseen.push(entry);
    } else if (stats.needsReview) {
      mistakes.push(entry);
    } else if (isDueForReview(stats)) {
      due.push(entry);
    } else {
      review.push(entry);
    }
  }

  shuffle(unseen);
  shuffle(mistakes);
  due.sort((first, second) =>
    getLastMasteredTime(progress[entryKey(first)]) - getLastMasteredTime(progress[entryKey(second)])
  );
  shuffle(review);

  const newTarget = Math.ceil(count / 2);
  const reviewTarget = count - newTarget;
  let mistakeTarget = reviewTarget;
  let dueTarget = 0;
  if (mistakes.length > 0 && due.length > 0) {
    mistakeTarget = Math.ceil(reviewTarget / 2);
    dueTarget = reviewTarget - mistakeTarget;
  } else if (due.length > 0) {
    mistakeTarget = 0;
    dueTarget = reviewTarget;
  }
  const selected = [];
  const selectedKeys = new Set();

  addQuestions(selected, selectedKeys, unseen.splice(0, newTarget), "new");
  addQuestions(selected, selectedKeys, mistakes.splice(0, mistakeTarget), "mistake");
  addQuestions(selected, selectedKeys, due.splice(0, dueTarget), "due");

  const fillPools = [
    { entries: mistakes, type: "mistake" },
    { entries: due, type: "due" },
    { entries: unseen, type: "new" },
    { entries: review, type: "review" }
  ];

  for (const pool of fillPools) {
    if (selected.length >= count) {
      break;
    }
    addQuestions(
      selected,
      selectedKeys,
      pool.entries.slice(0, count - selected.length),
      pool.type
    );
  }

  shuffle(selected);
  saveJson(DAILY_STORAGE_KEY, {
    date: today,
    count,
    items: selected.map(entry => ({
      key: entryKey(entry),
      type: entry.challengeType
    }))
  });
  return selected;
}

function addQuestions(target, selectedKeys, entries, challengeType) {
  for (const entry of entries) {
    const key = entryKey(entry);
    if (selectedKeys.has(key)) {
      continue;
    }
    selectedKeys.add(key);
    target.push({ ...entry, challengeType });
  }
}

function countQuestionTypes(questions) {
  return questions.reduce((counts, question) => {
    counts[question.challengeType || "review"] += 1;
    return counts;
  }, { new: 0, mistake: 0, due: 0, review: 0 });
}

function getPoolSummary(vocabulary) {
  const progress = loadProgress();
  return vocabulary.reduce((summary, entry) => {
    const stats = progress[entryKey(entry)];
    if (!stats || stats.seenCount === 0) {
      summary.unseen += 1;
    }
    if (stats?.needsReview) {
      summary.mistakes += 1;
    }
    if (stats && !stats.needsReview && isDueForReview(stats)) {
      summary.due += 1;
    }
    return summary;
  }, { unseen: 0, mistakes: 0, due: 0 });
}

function recordSeen(question) {
  updateProgress(question, stats => {
    stats.seenCount += 1;
    stats.lastSeen = new Date().toISOString();
  });
}

function recordMistake(question) {
  updateProgress(question, stats => {
    stats.wrongCount += 1;
    stats.needsReview = true;
    stats.lastWrong = new Date().toISOString();
  });
}

function recordCorrect(question, firstTry) {
  updateProgress(question, stats => {
    stats.correctCount += 1;
    stats.lastCorrect = new Date().toISOString();
    if (firstTry) {
      stats.needsReview = false;
      stats.lastMastered = stats.lastCorrect;
    }
  });
}

function isDueForReview(stats) {
  const masteredTime = getLastMasteredTime(stats);
  if (!masteredTime) {
    return false;
  }
  return Date.now() - masteredTime >= REVIEW_INTERVAL_DAYS * 24 * 60 * 60 * 1000;
}

function getLastMasteredTime(stats) {
  const value = stats?.lastMastered || stats?.lastCorrect;
  const timestamp = value ? Date.parse(value) : Number.NaN;
  return Number.isFinite(timestamp) ? timestamp : 0;
}

function updateProgress(question, updater) {
  const progress = loadProgress();
  const key = entryKey(question);
  const stats = progress[key] || {
    seenCount: 0,
    correctCount: 0,
    wrongCount: 0,
    needsReview: false,
    lastSeen: null,
    lastWrong: null,
    lastCorrect: null,
    lastMastered: null
  };
  updater(stats);
  progress[key] = stats;
  saveJson(PROGRESS_STORAGE_KEY, progress);
}

function loadProgress() {
  return loadJson(PROGRESS_STORAGE_KEY, {});
}

function entryKey(entry) {
  return `${normalizeAnswer(entry.english)}|${cleanText(entry.chinese)}`;
}

function getLocalDateKey() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function loadJson(key, fallback) {
  const value = window.localStorage.getItem(key);
  if (!value) {
    return fallback;
  }
  try {
    return JSON.parse(value);
  } catch (error) {
    console.error(`无法读取本地学习记录：${key}`, error);
    return fallback;
  }
}

function saveJson(key, value) {
  try {
    window.localStorage.setItem(key, JSON.stringify(value));
  } catch (error) {
    console.error(`无法保存本地学习记录：${key}`, error);
  }
}

function finishGame() {
  const answered = state.roundResults.length;
  const total = state.questions.length;
  const completedCorrect = state.correctCount;
  const accuracy = answered === 0 ? 0 : Math.round((completedCorrect / answered) * 100);
  const correctResults = state.roundResults.filter(result => result.outcome === "correct");
  const wrongResults = state.roundResults.filter(result => result.outcome === "wrong");

  elements.progressBar.style.width = "100%";
  elements.accuracyScore.textContent = `${accuracy}%`;
  elements.scoreGrade.textContent = getScoreGrade(accuracy);
  elements.correctResultCount.textContent = String(correctResults.length);
  elements.wrongResultCount.textContent = String(wrongResults.length);
  elements.resultSummary.textContent =
    `完成 ${answered} / ${total} 题；首次答对 ${completedCorrect} 题，错题 ${wrongResults.length} 题。`;
  elements.correctList.innerHTML = renderResultItems(correctResults, "本轮没有首次答对的词条。");
  elements.mistakeList.innerHTML = renderResultItems(wrongResults, "本轮没有错题，表现很好！");

  window.speechSynthesis?.cancel();
  showPanel(elements.resultPanel);
}

function renderResultItems(results, emptyMessage) {
  if (results.length === 0) {
    return `<p>${emptyMessage}</p>`;
  }
  return results.map(result =>
    `<div class="result-item"><span>${escapeHtml(result.question.chinese)}</span><span>${escapeHtml(result.question.english)}</span></div>`
  ).join("");
}

function getScoreGrade(accuracy) {
  if (accuracy >= 90) {
    return "优秀";
  }
  if (accuracy >= 80) {
    return "良好";
  }
  if (accuracy >= 60) {
    return "及格";
  }
  return "继续努力";
}

function resetToSetup() {
  window.speechSynthesis?.cancel();
  showPanel(elements.setupPanel);
}

function showPanel(panel) {
  for (const candidate of [elements.setupPanel, elements.gamePanel, elements.resultPanel]) {
    candidate.classList.toggle("hidden", candidate !== panel);
  }
}

function showLoadError(message) {
  state.vocabulary = [];
  elements.loadStatus.className = "status error";
  elements.loadStatus.textContent = message;
  elements.startButton.disabled = true;
}

function shuffle(items) {
  for (let index = items.length - 1; index > 0; index -= 1) {
    const randomIndex = Math.floor(Math.random() * (index + 1));
    [items[index], items[randomIndex]] = [items[randomIndex], items[index]];
  }
  return items;
}

function escapeHtml(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
