const $ = (selector, scope = document) => scope.querySelector(selector);
const $$ = (selector, scope = document) => Array.from(scope.querySelectorAll(selector));

const app = {
  activeTabLabel: $("#active-tab-label"),
  answerSummary: $("#answer-summary"),
  answerSteps: $("#answer-steps"),
  answerTrigger: $("#answer-trigger"),
  answerEvidence: $("#answer-evidence"),
  appPanels: $$("[data-app-panel]"),
  artifactBadge: $("#artifact-badge"),
  artifactList: $("#artifact-list"),
  artifactOutput: $("#artifact-output"),
  clearTrigger: $("#clear-trigger"),
  diagramOutput: $("#diagram-output"),
  difficulty: $("#difficulty"),
  downloadAll: $("#download-all"),
  downloadAnswer: $("#download-answer"),
  downloadCurrent: $("#download-current"),
  downloadGrade: $("#download-grade"),
  generateTrigger: $("#generate-trigger"),
  gradeConfidence: $("#grade-confidence"),
  gradeFile: $("#grading-file"),
  gradeLetter: $("#grade-letter"),
  gradeScore: $("#grade-score"),
  gradingSourceFiles: $("#grading-source-files"),
  gradingSourceText: $("#grading-source-text"),
  gradeTrigger: $("#grade-trigger"),
  gradingRubric: $("#grading-rubric"),
  gradingStatus: $("#grading-status"),
  gradingText: $("#grading-text"),
  generatorFiles: $("#generator-files"),
  generatorStatus: $("#generator-status"),
  generatorText: $("#generator-text"),
  ingestStatus: $("#ingest-status"),
  ingestTrigger: $("#ingest-trigger"),
  outputType: $("#output-type"),
  previewTitle: $("#preview-title"),
  printCurrent: $("#print-current"),
  qaInput: $("#qa-input"),
  qaSourceFiles: $("#qa-source-files"),
  qaSourceText: $("#qa-source-text"),
  qaStatus: $("#qa-status"),
  questionBankBadge: $("#question-bank-badge"),
  questionCount: $("#question-count"),
  questionList: $("#question-list"),
  refreshQuestions: $("#refresh-questions"),
  sampleTrigger: $("#sample-trigger"),
  scoreBreakdown: $("#score-breakdown"),
  sideTabs: $$("[data-app-tab]"),
  sourceConcepts: $("#source-concepts"),
  sourceCount: $("#source-count"),
  sourceFiles: $("#source-files"),
  sourceList: $("#source-list"),
  sourcePreview: $("#source-preview"),
  sourceQuestions: $("#source-questions"),
  sourceText: $("#source-text"),
  sourceWords: $("#source-words"),
  strengthsList: $("#strengths-list"),
  gapsList: $("#gaps-list"),
  feedbackSummary: $("#feedback-summary"),
  jumpTabs: $$("[data-jump-tab]"),
};

const state = {
  activeTab: "overview",
  artifacts: [],
  currentAnswer: null,
  currentArtifactId: null,
  currentGrade: null,
  knowledge: null,
  qaKnowledge: null,
  sources: [],
};

const SAMPLE_TEXT = `Photosynthesis Study Notes

Photosynthesis is the process plants use to convert light energy into chemical energy. It happens mainly inside the chloroplasts of plant cells. Chlorophyll absorbs light and helps start the reactions.

The process has two linked stages. In the light-dependent reactions, light energy is captured and used to produce ATP and NADPH. In the Calvin cycle, those products help build glucose from carbon dioxide.

Plants also need water for photosynthesis. Water is split during the light-dependent stage, which releases oxygen as a by-product. That oxygen leaves the plant and enters the atmosphere.

The overall word equation is carbon dioxide + water -> glucose + oxygen. Light is required because it provides the energy that starts the first stage and supports the second stage.

Why is chlorophyll important in photosynthesis?
How are the light-dependent reactions connected to the Calvin cycle?
What happens to oxygen during photosynthesis?
Explain why light is necessary for glucose production.`;

const STOPWORDS = new Set([
  "a",
  "about",
  "after",
  "again",
  "all",
  "also",
  "am",
  "an",
  "and",
  "any",
  "are",
  "as",
  "at",
  "be",
  "because",
  "been",
  "before",
  "being",
  "between",
  "both",
  "but",
  "by",
  "can",
  "could",
  "did",
  "do",
  "does",
  "doing",
  "down",
  "during",
  "each",
  "few",
  "for",
  "from",
  "further",
  "had",
  "has",
  "have",
  "having",
  "he",
  "her",
  "here",
  "hers",
  "herself",
  "him",
  "himself",
  "his",
  "how",
  "i",
  "if",
  "in",
  "into",
  "is",
  "it",
  "its",
  "itself",
  "just",
  "me",
  "more",
  "most",
  "my",
  "myself",
  "no",
  "nor",
  "not",
  "now",
  "of",
  "off",
  "on",
  "once",
  "only",
  "or",
  "other",
  "our",
  "ours",
  "ourselves",
  "out",
  "over",
  "own",
  "same",
  "she",
  "should",
  "so",
  "some",
  "such",
  "than",
  "that",
  "the",
  "their",
  "theirs",
  "them",
  "themselves",
  "then",
  "there",
  "these",
  "they",
  "this",
  "those",
  "through",
  "to",
  "too",
  "under",
  "until",
  "up",
  "very",
  "was",
  "we",
  "were",
  "what",
  "when",
  "where",
  "which",
  "while",
  "who",
  "why",
  "will",
  "with",
  "would",
  "you",
  "your",
  "yours",
  "yourself",
  "yourselves",
]);

const difficultyLabels = {
  challenge: "Challenge",
  foundation: "Foundation",
  standard: "Standard",
};

const escapeHtml = (value) =>
  String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");

const cleanText = (value) =>
  String(value || "")
    .replace(/\r/g, "\n")
    .replace(/\u00a0/g, " ")
    .replace(/[ \t]+/g, " ")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

const compact = (value, limit = 160) => {
  const normalized = cleanText(value).replace(/\n+/g, " ");

  if (normalized.length <= limit) {
    return normalized;
  }

  return `${normalized.slice(0, limit - 1).trimEnd()}…`;
};

const titleCase = (value) =>
  String(value || "")
    .split(/[\s-]+/)
    .filter(Boolean)
    .map((word) => `${word.charAt(0).toUpperCase()}${word.slice(1)}`)
    .join(" ");

const slugify = (value) =>
  String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 48);

const unique = (items) => [...new Set(items)];

const tokenize = (value) =>
  cleanText(value)
    .toLowerCase()
    .replace(/[^a-z0-9\s-]/g, " ")
    .split(/\s+/)
    .filter((token) => token.length > 2 && !STOPWORDS.has(token) && !/^\d+$/.test(token));

const sentenceSplit = (value) =>
  cleanText(value)
    .split(/(?<=[.!?])\s+|\n+/)
    .map((sentence) => sentence.trim())
    .filter((sentence) => sentence.length > 24);

const countWords = (value) => cleanText(value).split(/\s+/).filter(Boolean).length;

const setStatus = (element, message, tone = "idle") => {
  element.textContent = message;
  element.dataset.tone = tone;
};

const createSourceEntry = ({ label, mode, name, text, type }) => ({
  label,
  mode,
  name,
  text,
  type,
  wordCount: countWords(text),
});

const combineSourcesText = (sources) =>
  cleanText(
    sources
      .map((source) => `${source.name}\n${source.text}`)
      .join("\n\n")
  );

const setActiveTab = (tabName, { focusTab = false } = {}) => {
  state.activeTab = tabName;

  app.sideTabs.forEach((button) => {
    const isActive = button.dataset.appTab === tabName;

    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-selected", String(isActive));
    button.tabIndex = isActive ? 0 : -1;

    if (isActive && focusTab) {
      button.focus();
    }
  });

  app.appPanels.forEach((panel) => {
    const isActive = panel.dataset.appPanel === tabName;

    panel.classList.toggle("is-active", isActive);
    panel.hidden = !isActive;
  });

  const activeButton = app.sideTabs.find((button) => button.dataset.appTab === tabName);

  if (app.activeTabLabel && activeButton) {
    const label = $("strong", activeButton);
    app.activeTabLabel.textContent = label ? label.textContent : activeButton.textContent.trim();
  }
};

const downloadText = (fileName, text) => {
  const blob = new Blob([text], { type: "text/markdown;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
};

const fileExtension = (name = "") => {
  const parts = name.toLowerCase().split(".");
  return parts.length > 1 ? parts.pop() : "";
};

const stripRtf = (value) =>
  String(value || "")
    .replace(/\\par[d]?/g, "\n")
    .replace(/\\'[0-9a-f]{2}/gi, "")
    .replace(/\\[a-z]+\d* ?/gi, "")
    .replace(/[{}]/g, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

const scoreOverlap = (source, candidateTokens) => {
  if (!candidateTokens.length) {
    return 0;
  }

  const sourceTokens = unique(tokenize(source));
  let points = 0;

  candidateTokens.forEach((token) => {
    if (sourceTokens.includes(token)) {
      points += token.length > 7 ? 2 : 1;
    }
  });

  return points;
};

const jaccard = (left, right) => {
  const a = unique(tokenize(left));
  const b = unique(tokenize(right));

  if (!a.length || !b.length) {
    return 0;
  }

  const setB = new Set(b);
  const shared = a.filter((token) => setB.has(token)).length;
  const total = new Set([...a, ...b]).size;

  return shared / total;
};

const rankKeywords = (text) => {
  const counts = new Map();

  tokenize(text).forEach((token) => {
    counts.set(token, (counts.get(token) || 0) + 1);
  });

  return [...counts.entries()]
    .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0]))
    .map(([term, count]) => ({ term, count }));
};

const selectConcepts = (sentences, keywords) => {
  const topTokens = keywords.slice(0, 18).map((entry) => entry.term);
  const ranked = sentences
    .map((sentence) => {
      const overlap = scoreOverlap(sentence, topTokens);
      const lengthBonus = sentence.length > 60 && sentence.length < 220 ? 1.25 : 1;

      return {
        sentence,
        score: overlap * lengthBonus,
      };
    })
    .sort((left, right) => right.score - left.score || left.sentence.length - right.sentence.length);

  const picked = [];

  ranked.forEach((entry) => {
    if (picked.length >= 10) {
      return;
    }

    const isNearDuplicate = picked.some((existing) => jaccard(existing, entry.sentence) > 0.72);

    if (!isNearDuplicate && entry.score > 0) {
      picked.push(entry.sentence);
    }
  });

  return picked.length ? picked : sentences.slice(0, 8);
};

const buildTerms = (keywords, sentences) =>
  keywords.slice(0, 12).map((entry) => {
    const sentence =
      sentences.find((candidate) => candidate.toLowerCase().includes(entry.term)) ||
      sentences[0] ||
      "";

    return {
      term: entry.term,
      explanation: compact(sentence, 160),
    };
  });

const extractQuestions = (rawText, terms, concepts) => {
  const explicit = [];

  rawText
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean)
    .forEach((line) => {
      const matches = line.match(/[^?]+\?/g);

      if (matches) {
        matches.forEach((match) => explicit.push(cleanText(match)));
      }
    });

  const generated = [
    ...terms.slice(0, 4).map(({ term }) => `What does ${term} mean in this material?`),
    ...terms
      .slice(4, 8)
      .map(({ term }) => `How would you explain the role of ${term} using the source?`),
    ...concepts
      .slice(0, 4)
      .map((sentence) => `Which idea in the source best explains: ${compact(sentence, 70)}?`),
  ];

  return unique([...explicit, ...generated]).slice(0, 12);
};

const deriveTitle = (rawText, keywords) => {
  const firstLine = rawText
    .split(/\n+/)
    .map((line) => line.trim())
    .find((line) => line.length > 8 && line.length < 80);

  if (firstLine) {
    return firstLine.replace(/[#:*]/g, "");
  }

  if (keywords.length >= 2) {
    return `${titleCase(keywords[0].term)} and ${titleCase(keywords[1].term)}`;
  }

  return "Study Pack";
};

const buildKnowledgeModel = (rawText) => {
  const normalized = cleanText(rawText);
  const paragraphs = normalized.split(/\n{2,}/).filter(Boolean);
  const sentences = sentenceSplit(normalized);
  const keywords = rankKeywords(normalized);
  const concepts = selectConcepts(sentences, keywords);
  const terms = buildTerms(keywords, sentences);
  const questions = extractQuestions(rawText, terms, concepts);

  return {
    concepts,
    normalized,
    paragraphs,
    questions,
    sentences,
    stats: {
      conceptCount: concepts.length,
      questionCount: questions.length,
      wordCount: countWords(normalized),
    },
    terms,
    title: deriveTitle(rawText, keywords),
    topKeywords: keywords,
    overview: concepts.slice(0, 3),
  };
};

const buildKnowledgeFromSources = (sources) => {
  const combinedText = combineSourcesText(sources);

  if (!combinedText) {
    return null;
  }

  const knowledge = buildKnowledgeModel(combinedText);
  knowledge.sourceCount = sources.length;
  return knowledge;
};

const readSourceInputs = async ({
  files = [],
  statusElement = null,
  text = "",
  textName = "Manual source",
}) => {
  const sources = [];
  const errors = [];
  const preparedText = cleanText(text);

  if (preparedText) {
    sources.push(
      createSourceEntry({
        label: "Pasted text",
        mode: "Pasted text",
        name: textName,
        text: preparedText,
        type: "text",
      })
    );
  }

  for (const file of files) {
    try {
      if (statusElement) {
        setStatus(statusElement, `Reading ${file.name}...`, "working");
      }

      const extractedText = await extractTextFromFile(file);

      if (!extractedText) {
        throw new Error("No readable text was found in that file.");
      }

      sources.push(
        createSourceEntry({
          label: buildSourceLabel({
            mode: "Uploaded file",
            type: file.type || fileExtension(file.name) || "unknown",
          }),
          mode: "Uploaded file",
          name: file.name,
          text: extractedText,
          type: file.type || fileExtension(file.name) || "unknown",
        })
      );
    } catch (error) {
      console.error(error);
      errors.push({
        fileName: file.name,
        message: error.message || "Please try another format.",
      });
    }
  }

  return { errors, sources };
};

const resolveKnowledgeFromInputs = async ({
  emptyMessage,
  fallbackKnowledge = state.knowledge,
  fallbackSources = state.sources,
  files = [],
  statusElement = null,
  text = "",
  textName = "Manual source",
}) => {
  const { errors, sources } = await readSourceInputs({
    files,
    statusElement,
    text,
    textName,
  });

  if (sources.length) {
    const knowledge = buildKnowledgeFromSources(sources);

    if (!knowledge) {
      if (statusElement) {
        setStatus(statusElement, "No readable text was extracted from the provided source.", "error");
      }

      return null;
    }

    return {
      errors,
      knowledge,
      origin: "direct",
      sources,
    };
  }

  if (fallbackKnowledge) {
    return {
      errors,
      knowledge: fallbackKnowledge,
      origin: "workspace",
      sources: fallbackSources,
    };
  }

  if (statusElement && emptyMessage) {
    setStatus(statusElement, emptyMessage, "error");
  }

  return null;
};

const markdownToHtml = (markdown) => {
  const lines = String(markdown || "").split("\n");
  const html = [];
  let listType = null;

  const closeList = () => {
    if (listType) {
      html.push(`</${listType}>`);
      listType = null;
    }
  };

  lines.forEach((line) => {
    const trimmed = line.trim();

    if (!trimmed) {
      closeList();
      return;
    }

    if (trimmed.startsWith("### ")) {
      closeList();
      html.push(`<h3>${escapeHtml(trimmed.slice(4))}</h3>`);
      return;
    }

    if (trimmed.startsWith("## ")) {
      closeList();
      html.push(`<h2>${escapeHtml(trimmed.slice(3))}</h2>`);
      return;
    }

    if (trimmed.startsWith("# ")) {
      closeList();
      html.push(`<h1>${escapeHtml(trimmed.slice(2))}</h1>`);
      return;
    }

    if (trimmed.startsWith("> ")) {
      closeList();
      html.push(`<blockquote>${escapeHtml(trimmed.slice(2))}</blockquote>`);
      return;
    }

    if (/^- /.test(trimmed)) {
      if (listType !== "ul") {
        closeList();
        listType = "ul";
        html.push("<ul>");
      }

      html.push(`<li>${escapeHtml(trimmed.slice(2))}</li>`);
      return;
    }

    if (/^\d+\.\s/.test(trimmed)) {
      if (listType !== "ol") {
        closeList();
        listType = "ol";
        html.push("<ol>");
      }

      html.push(`<li>${escapeHtml(trimmed.replace(/^\d+\.\s/, ""))}</li>`);
      return;
    }

    closeList();
    html.push(`<p>${escapeHtml(trimmed)}</p>`);
  });

  closeList();

  return html.join("");
};

const createMcqSet = (knowledge, count) => {
  const pool = knowledge.terms.length >= 4 ? knowledge.terms : knowledge.topKeywords.slice(0, 8);
  const questions = [];

  for (let index = 0; index < count; index += 1) {
    const focus = pool[index % pool.length];
    const distractors = pool
      .filter((entry) => entry.term !== focus.term)
      .slice(index % Math.max(pool.length - 1, 1))
      .concat(pool.filter((entry) => entry.term !== focus.term))
      .slice(0, 3)
      .map((entry) => titleCase(entry.term));
    const options = unique([titleCase(focus.term), ...distractors]).slice(0, 4);
    const answer = titleCase(focus.term);

    questions.push({
      answer,
      explanation: focus.explanation,
      options,
      prompt: `Which term best matches this clue from the source: "${focus.explanation}"`,
    });
  }

  return questions;
};

const createShortAnswerSet = (knowledge, count) =>
  knowledge.concepts.slice(0, count).map((concept, index) => ({
    answer: concept,
    prompt: `${index + 1}. Explain this core idea in your own words: ${compact(concept, 96)}`,
  }));

const createAppliedQuestions = (knowledge, count) =>
  knowledge.questions.slice(0, count).map((question, index) => ({
    answer: compact(
      knowledge.concepts[index % knowledge.concepts.length] || knowledge.overview[0] || "",
      180
    ),
    prompt: `${index + 1}. ${question}`,
  }));

const buildStudyGuide = (knowledge, questionCount, difficulty) => {
  const practiceCount = Math.min(questionCount, 8);
  const sourceCount = knowledge.sourceCount || state.sources.length || 1;
  const markdown = [
    `# ${knowledge.title} Study Guide`,
    `> Difficulty: ${difficultyLabels[difficulty]} | Sources loaded: ${sourceCount}`,
    "## Snapshot",
    ...knowledge.overview.map((sentence) => `- ${sentence}`),
    "## Key Ideas",
    ...knowledge.concepts.slice(0, 6).map((sentence) => `- ${sentence}`),
    "## Vocabulary and Concepts",
    ...knowledge.terms.slice(0, 8).map((entry) => `- ${titleCase(entry.term)}: ${entry.explanation}`),
    "## Study Steps",
    "1. Read the overview once and underline the repeated ideas.",
    `2. Focus on these high-value terms: ${knowledge.terms
      .slice(0, 5)
      .map((entry) => titleCase(entry.term))
      .join(", ")}.`,
    "3. Turn each key idea into a short spoken explanation from memory.",
    "4. Use the self-check questions below before moving to the quiz or test.",
    "## Self-Check Questions",
    ...knowledge.questions.slice(0, practiceCount).map((question) => `- ${question}`),
  ].join("\n");

  return {
    fileName: `${slugify(knowledge.title)}-study-guide.md`,
    id: `artifact-${Date.now()}-study-guide`,
    markdown,
    title: "Study Guide",
    type: "study-guide",
  };
};

const buildQuiz = (knowledge, questionCount, difficulty) => {
  const mcqCount = Math.max(4, Math.floor(questionCount / 2));
  const shortCount = Math.max(2, questionCount - mcqCount);
  const mcqs = createMcqSet(knowledge, mcqCount);
  const shorts = createShortAnswerSet(knowledge, shortCount);
  const difficultyNote =
    difficulty === "foundation"
      ? "Keep answers direct and vocabulary-focused."
      : difficulty === "challenge"
        ? "Push for precise cause-and-effect explanations."
        : "Balance core ideas with clear explanations.";

  const lines = [
    `# ${knowledge.title} Quiz`,
    `> Difficulty: ${difficultyLabels[difficulty]} | ${difficultyNote}`,
    "## Multiple Choice",
  ];

  mcqs.forEach((question, index) => {
    lines.push(`${index + 1}. ${question.prompt}`);
    question.options.forEach((option, optionIndex) => {
      lines.push(`- ${String.fromCharCode(65 + optionIndex)}. ${option}`);
    });
  });

  lines.push("## Short Answer");

  shorts.forEach((question) => {
    lines.push(question.prompt);
  });

  lines.push("## Answer Key");

  mcqs.forEach((question, index) => {
    lines.push(`${index + 1}. ${question.answer} - ${question.explanation}`);
  });

  shorts.forEach((question, index) => {
    lines.push(`${mcqs.length + index + 1}. ${question.answer}`);
  });

  return {
    fileName: `${slugify(knowledge.title)}-quiz.md`,
    id: `artifact-${Date.now()}-quiz`,
    markdown: lines.join("\n"),
    title: "Quiz",
    type: "quiz",
  };
};

const buildTest = (knowledge, questionCount, difficulty) => {
  const mcqCount = Math.max(5, Math.ceil(questionCount * 0.45));
  const shortCount = Math.max(3, Math.floor(questionCount * 0.35));
  const appliedCount = Math.max(2, questionCount - mcqCount - shortCount);
  const mcqs = createMcqSet(knowledge, mcqCount);
  const shorts = createShortAnswerSet(knowledge, shortCount);
  const applied = createAppliedQuestions(knowledge, appliedCount);
  const sourceCount = knowledge.sourceCount || state.sources.length || 1;

  const lines = [
    `# ${knowledge.title} Practice Test`,
    `> Difficulty: ${difficultyLabels[difficulty]} | Built from ${sourceCount} imported sources`,
    "## Section A: Multiple Choice",
  ];

  mcqs.forEach((question, index) => {
    lines.push(`${index + 1}. ${question.prompt}`);
    question.options.forEach((option, optionIndex) => {
      lines.push(`- ${String.fromCharCode(65 + optionIndex)}. ${option}`);
    });
  });

  lines.push("## Section B: Short Response");

  shorts.forEach((question) => {
    lines.push(question.prompt);
  });

  lines.push("## Section C: Extended Response");

  applied.forEach((question) => {
    lines.push(question.prompt);
  });

  lines.push("## Scoring Guide");

  mcqs.forEach((question, index) => {
    lines.push(`${index + 1}. ${question.answer}`);
  });

  shorts.forEach((question, index) => {
    lines.push(`${mcqs.length + index + 1}. ${question.answer}`);
  });

  applied.forEach((question, index) => {
    lines.push(
      `${mcqs.length + shorts.length + index + 1}. Strong responses should connect ideas like: ${question.answer}`
    );
  });

  return {
    fileName: `${slugify(knowledge.title)}-practice-test.md`,
    id: `artifact-${Date.now()}-test`,
    markdown: lines.join("\n"),
    title: "Practice Test",
    type: "test",
  };
};

const buildArtifact = (knowledge, type, questionCount, difficulty) => {
  if (!knowledge) {
    return null;
  }

  if (type === "quiz") {
    return buildQuiz(knowledge, questionCount, difficulty);
  }

  if (type === "test") {
    return buildTest(knowledge, questionCount, difficulty);
  }

  return buildStudyGuide(knowledge, questionCount, difficulty);
};

const getCurrentArtifact = () =>
  state.artifacts.find((artifact) => artifact.id === state.currentArtifactId) || null;

const renderSourceList = () => {
  if (!state.sources.length) {
    app.sourceList.innerHTML = `
      <div class="empty-state">
        Import material to see parsed files, extracted OCR text, and source coverage.
      </div>
    `;
    return;
  }

  app.sourceList.innerHTML = state.sources
    .map(
      (source) => `
        <article class="source-item">
          <div>
            <strong>${escapeHtml(source.name)}</strong>
            <div class="source-meta">${escapeHtml(source.label)}</div>
          </div>
          <div class="source-meta">${source.wordCount} words extracted</div>
        </article>
      `
    )
    .join("");
};

const renderSourceOverview = () => {
  if (!state.knowledge) {
    app.sourceCount.textContent = "0";
    app.sourceWords.textContent = "0";
    app.sourceQuestions.textContent = "0";
    app.sourceConcepts.textContent = "0";
    app.previewTitle.textContent = "No source loaded";
    app.sourcePreview.textContent = "Your combined study text will appear here after ingestion.";
    return;
  }

  app.sourceCount.textContent = String(state.sources.length);
  app.sourceWords.textContent = String(state.knowledge.stats.wordCount);
  app.sourceQuestions.textContent = String(state.knowledge.stats.questionCount);
  app.sourceConcepts.textContent = String(state.knowledge.stats.conceptCount);
  app.previewTitle.textContent = state.knowledge.title;
  app.sourcePreview.textContent = compact(state.knowledge.normalized, 1200);
};

const renderArtifactOutput = () => {
  const artifact = getCurrentArtifact();

  if (!artifact) {
    app.artifactBadge.textContent = "Nothing generated yet";
    app.artifactOutput.innerHTML = `
      <div class="empty-state">
        Generate a study guide, quiz, or practice test to preview it here.
      </div>
    `;
    app.downloadCurrent.disabled = true;
    app.downloadAll.disabled = state.artifacts.length === 0;
    app.printCurrent.disabled = true;
    return;
  }

  app.artifactBadge.textContent = artifact.title;
  app.artifactOutput.innerHTML = `
    <div class="rendered-markdown">${markdownToHtml(artifact.markdown)}</div>
  `;
  app.downloadCurrent.disabled = false;
  app.downloadAll.disabled = state.artifacts.length === 0;
  app.printCurrent.disabled = false;
};

const renderArtifactList = () => {
  if (!state.artifacts.length) {
    app.artifactList.innerHTML = `
      <div class="empty-state">
        Your generated study guide, quiz, and test files will collect here.
      </div>
    `;
    renderArtifactOutput();
    return;
  }

  app.artifactList.innerHTML = state.artifacts
    .map(
      (artifact) => `
        <article class="artifact-item${artifact.id === state.currentArtifactId ? " is-active" : ""}">
          <div>
            <strong>${escapeHtml(artifact.title)}</strong>
            <div class="artifact-meta">${escapeHtml(artifact.fileName)}</div>
          </div>
          <div class="artifact-actions">
            <button class="button secondary" type="button" data-artifact-open="${artifact.id}">
              Open
            </button>
            <button class="button ghost" type="button" data-artifact-download="${artifact.id}">
              Download
            </button>
          </div>
        </article>
      `
    )
    .join("");

  $$("[data-artifact-open]").forEach((button) => {
    button.addEventListener("click", () => {
      state.currentArtifactId = button.dataset.artifactOpen;
      renderArtifactList();
      renderArtifactOutput();
    });
  });

  $$("[data-artifact-download]").forEach((button) => {
    button.addEventListener("click", () => {
      const artifact = state.artifacts.find((entry) => entry.id === button.dataset.artifactDownload);

      if (artifact) {
        downloadText(artifact.fileName, artifact.markdown);
      }
    });
  });

  renderArtifactOutput();
};

const renderQuestionList = (knowledge = state.qaKnowledge || state.knowledge) => {
  const questions = knowledge ? knowledge.questions : [];
  app.questionBankBadge.textContent = `${questions.length} question${questions.length === 1 ? "" : "s"}`;

  if (!questions.length) {
    app.questionList.innerHTML = `
      <div class="empty-state">
        Build a study pack first, then this tab will surface file questions automatically.
      </div>
    `;
    return;
  }

  app.questionList.innerHTML = questions
    .map(
      (question, index) => `
        <article class="question-item">
          <div>
            <strong>Question ${index + 1}</strong>
            <small>${escapeHtml(question)}</small>
          </div>
          <button class="button secondary" type="button" data-question="${escapeHtml(question)}">
            Use Question
          </button>
        </article>
      `
    )
    .join("");

  $$("[data-question]").forEach((button) => {
    button.addEventListener("click", () => {
      app.qaInput.value = button.dataset.question;
      handleAnswerQuestion();
    });
  });
};

const renderAnswer = async () => {
  if (!state.currentAnswer) {
    app.answerSummary.textContent =
      "The answer summary will appear here once a question is processed.";
    app.answerSteps.innerHTML = '<div class="empty-state">No explanation generated yet.</div>';
    app.answerEvidence.innerHTML =
      '<div class="empty-state">Relevant source lines will appear here.</div>';
    app.diagramOutput.innerHTML =
      '<div class="empty-state">A source-based diagram will render here.</div>';
    app.downloadAnswer.disabled = true;
    return;
  }

  app.answerSummary.textContent = state.currentAnswer.summary;
  app.answerSteps.innerHTML = state.currentAnswer.steps
    .map(
      (step, index) => `
        <article class="step-item">
          <strong>Step ${index + 1}: ${escapeHtml(step.title)}</strong>
          <div>${escapeHtml(step.body)}</div>
        </article>
      `
    )
    .join("");
  app.answerEvidence.innerHTML = state.currentAnswer.evidence
    .map(
      (sentence, index) => `
        <article class="evidence-item">
          <strong>Evidence ${index + 1}</strong>
          <small>${escapeHtml(sentence)}</small>
        </article>
      `
    )
    .join("");

  app.downloadAnswer.disabled = false;
  await renderDiagram(state.currentAnswer.diagram);
};

const renderGrade = () => {
  if (!state.currentGrade) {
    app.gradeScore.textContent = "0%";
    app.gradeLetter.textContent = "-";
    app.gradeConfidence.textContent = "0%";
    app.scoreBreakdown.innerHTML = '<div class="empty-state">No grading breakdown yet.</div>';
    app.strengthsList.innerHTML =
      '<div class="empty-state">Strengths will appear after grading.</div>';
    app.gapsList.innerHTML =
      '<div class="empty-state">Missing ideas and weak areas will appear here.</div>';
    app.feedbackSummary.textContent = "The grading summary will appear here after a scan.";
    app.downloadGrade.disabled = true;
    return;
  }

  app.gradeScore.textContent = `${state.currentGrade.score}%`;
  app.gradeLetter.textContent = state.currentGrade.letter;
  app.gradeConfidence.textContent = `${state.currentGrade.confidence}%`;
  app.scoreBreakdown.innerHTML = state.currentGrade.breakdown
    .map(
      (item) => `
        <article class="breakdown-row">
          <strong>${escapeHtml(item.label)}</strong>
          <span>${item.value}%</span>
        </article>
      `
    )
    .join("");
  app.strengthsList.innerHTML = state.currentGrade.strengths
    .map((item) => `<article class="feedback-item">${escapeHtml(item)}</article>`)
    .join("");
  app.gapsList.innerHTML = state.currentGrade.gaps
    .map((item) => `<article class="feedback-item">${escapeHtml(item)}</article>`)
    .join("");
  app.feedbackSummary.textContent = state.currentGrade.summary;
  app.downloadGrade.disabled = false;
};

const resetArtifacts = () => {
  state.artifacts = [];
  state.currentArtifactId = null;
  renderArtifactList();
};

const resetAnswer = () => {
  state.currentAnswer = null;
  renderAnswer();
};

const resetGrade = () => {
  state.currentGrade = null;
  renderGrade();
};

const resetWorkspace = () => {
  state.sources = [];
  state.knowledge = null;
  state.qaKnowledge = null;
  app.sourceFiles.value = "";
  app.sourceText.value = "";
  renderSourceList();
  renderSourceOverview();
  renderQuestionList();
  resetArtifacts();
  resetAnswer();
  resetGrade();
  setStatus(app.ingestStatus, "Waiting for source material.", "idle");
  setStatus(
    app.generatorStatus,
    "Using the current Workspace source unless you provide Generator files or text.",
    "idle"
  );
  setStatus(
    app.qaStatus,
    "Using the current Workspace source unless you provide Q&A files or text.",
    "idle"
  );
};

const extractTextFromPdf = async (file) => {
  if (!window.pdfjsLib) {
    throw new Error("PDF support is unavailable right now.");
  }

  const arrayBuffer = await file.arrayBuffer();
  const pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const pages = [];
  const pageCount = Math.min(pdf.numPages, 20);

  for (let index = 1; index <= pageCount; index += 1) {
    const page = await pdf.getPage(index);
    const textContent = await page.getTextContent();
    const pageText = textContent.items.map((item) => item.str).join(" ");
    pages.push(pageText);
  }

  return cleanText(pages.join("\n"));
};

const extractTextFromDocx = async (file) => {
  if (!window.mammoth) {
    throw new Error("DOCX support is unavailable right now.");
  }

  const arrayBuffer = await file.arrayBuffer();
  const result = await window.mammoth.extractRawText({ arrayBuffer });
  return cleanText(result.value);
};

const extractTextFromImage = async (file) => {
  if (!window.Tesseract) {
    throw new Error("OCR support is unavailable right now.");
  }

  const result = await window.Tesseract.recognize(file, "eng");
  return cleanText(result.data.text);
};

const extractTextFromFile = async (file) => {
  const extension = fileExtension(file.name);

  if (file.type.startsWith("image/")) {
    return extractTextFromImage(file);
  }

  if (extension === "pdf") {
    return extractTextFromPdf(file);
  }

  if (extension === "docx") {
    return extractTextFromDocx(file);
  }

  const text = await file.text();

  if (extension === "rtf") {
    return stripRtf(text);
  }

  return cleanText(text);
};

const buildSourceLabel = (source) =>
  `${source.mode}${source.mode === "Uploaded file" ? ` · ${source.type}` : ""}`;

const handleIngest = async ({ useSample = false } = {}) => {
  const files = Array.from(app.sourceFiles.files || []);
  const pastedText = useSample ? SAMPLE_TEXT : app.sourceText.value;
  const preparedText = cleanText(pastedText);

  if (!preparedText && !files.length) {
    setStatus(app.ingestStatus, "Add pasted text or upload at least one file first.", "error");
    return;
  }

  state.sources = [];
  state.knowledge = null;
  resetArtifacts();
  resetAnswer();
  resetGrade();
  renderSourceList();
  renderSourceOverview();
  renderQuestionList();
  setStatus(app.ingestStatus, "Building your study pack from the provided material...", "working");

  if (useSample) {
    app.sourceText.value = SAMPLE_TEXT;
  }

  const { errors, sources } = await readSourceInputs({
    files,
    statusElement: app.ingestStatus,
    text: preparedText,
    textName: "Manual source",
  });

  state.sources = sources;

  const knowledge = buildKnowledgeFromSources(state.sources);

  if (!knowledge) {
    setStatus(
      app.ingestStatus,
      errors.length
        ? `No readable text was extracted. Last issue: ${errors[errors.length - 1].message}`
        : "No readable text was extracted from the material.",
      "error"
    );
    renderSourceList();
    return;
  }

  state.knowledge = knowledge;
  state.qaKnowledge = null;
  renderSourceList();
  renderSourceOverview();
  renderQuestionList();
  setStatus(
    app.ingestStatus,
    `Study pack ready. ${state.knowledge.stats.wordCount} words parsed from ${state.sources.length} source${state.sources.length === 1 ? "" : "s"}.`,
    "ready"
  );

  await handleGenerateArtifact({ silent: true });
};

const handleGenerateArtifact = async ({ silent = false } = {}) => {
  const sourceResult = await resolveKnowledgeFromInputs({
    emptyMessage:
      "Add Generator files/text or build a Workspace study pack before generating outputs.",
    files: Array.from(app.generatorFiles.files || []),
    statusElement: app.generatorStatus,
    text: app.generatorText.value,
    textName: "Generator source",
  });

  if (!sourceResult) {
    return;
  }

  const questionCount = Number(app.questionCount.value || 12);
  const artifact = buildArtifact(
    sourceResult.knowledge,
    app.outputType.value,
    questionCount,
    app.difficulty.value
  );

  if (!artifact) {
    return;
  }

  state.artifacts.unshift(artifact);
  state.currentArtifactId = artifact.id;
  renderArtifactList();

  if (!silent) {
    const sourceLabel =
      sourceResult.origin === "direct"
        ? "from Generator files/text"
        : "from the Workspace study pack";
    setStatus(app.generatorStatus, `${artifact.title} generated ${sourceLabel}.`, "ready");
  }
};

const summarizeEvidence = (evidence) => {
  if (!evidence.length) {
    return "The source does not contain a close match, so review the key ideas and terms first.";
  }

  if (evidence.length === 1) {
    return evidence[0];
  }

  return `${evidence[0]} ${evidence[1]}`;
};

const buildMermaidDiagram = (question, evidence, summary) => {
  const cleanNode = (value, limit = 52) =>
    compact(value, limit).replaceAll('"', "'").replaceAll("[", "(").replaceAll("]", ")");

  return `flowchart LR
    A["Question: ${cleanNode(question, 44)}"] --> B["Evidence 1: ${cleanNode(
      evidence[0] || "Review the closest source statement."
    )}"]
    B --> C["Evidence 2: ${cleanNode(
      evidence[1] || evidence[0] || "Connect the idea to the source."
    )}"]
    C --> D["Answer: ${cleanNode(summary, 46)}"]`;
};

const generateAnswer = (question, knowledge) => {
  const trimmedQuestion = cleanText(question);

  if (!trimmedQuestion || !knowledge) {
    return null;
  }

  const tokens = unique(tokenize(trimmedQuestion));
  const ranked = knowledge.sentences
    .map((sentence) => ({
      score:
        scoreOverlap(sentence, tokens) +
        (sentence.toLowerCase().includes(trimmedQuestion.toLowerCase()) ? 4 : 0),
      sentence,
    }))
    .sort((left, right) => right.score - left.score);

  const evidence = ranked
    .filter((entry) => entry.score > 0)
    .slice(0, 3)
    .map((entry) => entry.sentence);

  if (!evidence.length) {
    evidence.push(...knowledge.overview.slice(0, 2));
  }

  const focusPhrase =
    tokens.slice(0, 4).map(titleCase).join(", ") || "the main idea from the source";
  const summary = `The best source-based answer is that ${summarizeEvidence(evidence)
    .replace(/^[A-Z]/, (match) => match.toLowerCase())} This answer focuses on ${focusPhrase}.`;
  const steps = [
    {
      body: `The question is mainly asking about ${focusPhrase}.`,
      title: "Find the focus",
    },
    {
      body: evidence[0] || "Review the closest sentence in the imported material.",
      title: "Pull the strongest clue",
    },
    {
      body: evidence[1]
        ? `Connect that clue to this supporting idea: ${evidence[1]}`
        : "Connect the clue back to the source summary so the explanation stays grounded.",
      title: "Add supporting evidence",
    },
    {
      body: summary,
      title: "State the answer clearly",
    },
  ];

  return {
    diagram: buildMermaidDiagram(trimmedQuestion, evidence, summary),
    evidence,
    markdown: [
      `# Answer Breakdown`,
      `## Question`,
      trimmedQuestion,
      `## Best Answer`,
      summary,
      `## Step-by-Step Explanation`,
      ...steps.map((step, index) => `${index + 1}. ${step.title}: ${step.body}`),
      `## Evidence`,
      ...evidence.map((item) => `- ${item}`),
    ].join("\n"),
    question: trimmedQuestion,
    steps,
    summary,
  };
};

const renderDiagram = async (definition) => {
  if (!definition) {
    app.diagramOutput.innerHTML =
      '<div class="empty-state">A source-based diagram will render here.</div>';
    return;
  }

  if (!window.mermaid) {
    app.diagramOutput.innerHTML = `<div class="empty-state">${escapeHtml(definition)}</div>`;
    return;
  }

  try {
    const { svg } = await window.mermaid.render(`diagram-${Date.now()}`, definition);
    app.diagramOutput.innerHTML = svg;
  } catch (error) {
    console.error(error);
    app.diagramOutput.innerHTML =
      '<div class="empty-state">The diagram could not be rendered, but the answer text is still available.</div>';
  }
};

const handleRefreshQuestions = async () => {
  const sourceResult = await resolveKnowledgeFromInputs({
    emptyMessage:
      "Add Q&A files/text or build a Workspace study pack before refreshing file questions.",
    files: Array.from(app.qaSourceFiles.files || []),
    statusElement: app.qaStatus,
    text: app.qaSourceText.value,
    textName: "Q&A source",
  });

  if (!sourceResult) {
    return;
  }

  state.qaKnowledge = sourceResult.origin === "direct" ? sourceResult.knowledge : null;
  renderQuestionList(sourceResult.knowledge);

  const sourceLabel =
    sourceResult.origin === "direct" ? "Q&A files/text" : "the Workspace study pack";
  setStatus(
    app.qaStatus,
    `${sourceResult.knowledge.questions.length} file questions loaded from ${sourceLabel}.`,
    "ready"
  );
};

const handleAnswerQuestion = async () => {
  const sourceResult = await resolveKnowledgeFromInputs({
    emptyMessage:
      "Add Q&A files/text or build a Workspace study pack before answering source questions.",
    files: Array.from(app.qaSourceFiles.files || []),
    statusElement: app.qaStatus,
    text: app.qaSourceText.value,
    textName: "Q&A source",
  });

  if (!sourceResult) {
    return;
  }

  state.qaKnowledge = sourceResult.origin === "direct" ? sourceResult.knowledge : null;
  renderQuestionList(sourceResult.knowledge);

  const answer = generateAnswer(app.qaInput.value, sourceResult.knowledge);

  if (!answer) {
    setStatus(app.qaStatus, "Enter or select a question before generating an answer.", "error");
    return;
  }

  state.currentAnswer = answer;
  await renderAnswer();
  const sourceLabel =
    sourceResult.origin === "direct" ? "Q&A files/text" : "the Workspace study pack";
  setStatus(app.qaStatus, `Single-question answer generated from ${sourceLabel}.`, "ready");
};

const getRubricWeights = (rubricText) => {
  const lowered = rubricText.toLowerCase();
  const weights = {
    accuracy: 0.35,
    coverage: 0.35,
    reasoning: 0.15,
    structure: 0.15,
  };

  if (lowered.includes("vocabulary")) {
    weights.coverage += 0.1;
    weights.structure -= 0.05;
    weights.reasoning -= 0.05;
  }

  if (lowered.includes("organization") || lowered.includes("structure")) {
    weights.structure += 0.1;
    weights.coverage -= 0.05;
    weights.reasoning -= 0.05;
  }

  if (lowered.includes("evidence")) {
    weights.reasoning += 0.1;
    weights.structure -= 0.05;
    weights.coverage -= 0.05;
  }

  if (lowered.includes("accuracy")) {
    weights.accuracy += 0.1;
    weights.reasoning -= 0.05;
    weights.structure -= 0.05;
  }

  return weights;
};

const scoreCoverage = (submissionText, keywords) => {
  const lowered = submissionText.toLowerCase();
  const focusTerms = keywords.slice(0, 14).map((entry) => entry.term);
  const matched = focusTerms.filter((term) => lowered.includes(term));

  return {
    present: matched,
    score: Math.round((matched.length / Math.max(focusTerms.length, 1)) * 100),
  };
};

const scoreAccuracy = (submissionText, concepts) => {
  const submissionSentences = sentenceSplit(submissionText);

  if (!submissionSentences.length || !concepts.length) {
    return 0;
  }

  const matches = concepts.slice(0, 6).map((concept) => {
    const best = submissionSentences.reduce((highest, sentence) => {
      const similarity = jaccard(concept, sentence);
      return Math.max(highest, similarity);
    }, 0);

    return best;
  });

  const average = matches.reduce((sum, value) => sum + value, 0) / matches.length;
  return Math.round(average * 100);
};

const scoreReasoning = (submissionText) => {
  const lowered = submissionText.toLowerCase();
  const markers = [
    "because",
    "therefore",
    "as a result",
    "for example",
    "first",
    "next",
    "finally",
    "so that",
    "which means",
  ];
  const hits = markers.filter((marker) => lowered.includes(marker)).length;
  const sentenceCount = sentenceSplit(submissionText).length;
  const lengthBoost = Math.min(25, Math.round(countWords(submissionText) / 8));

  return Math.min(100, hits * 12 + sentenceCount * 5 + lengthBoost);
};

const scoreStructure = (submissionText) => {
  const paragraphCount = cleanText(submissionText).split(/\n{2,}/).filter(Boolean).length;
  const sentenceCount = sentenceSplit(submissionText).length;
  const numberedAnswers = (submissionText.match(/\b\d+[.)]/g) || []).length;

  return Math.min(100, 35 + paragraphCount * 12 + sentenceCount * 4 + numberedAnswers * 5);
};

const letterGradeFor = (score) => {
  if (score >= 97) return "A+";
  if (score >= 93) return "A";
  if (score >= 90) return "A-";
  if (score >= 87) return "B+";
  if (score >= 83) return "B";
  if (score >= 80) return "B-";
  if (score >= 77) return "C+";
  if (score >= 73) return "C";
  if (score >= 70) return "C-";
  if (score >= 67) return "D+";
  if (score >= 63) return "D";
  if (score >= 60) return "D-";
  return "F";
};

const buildGradeReport = (submissionText, rubricText, knowledge) => {
  const weights = getRubricWeights(rubricText);
  const coverage = scoreCoverage(submissionText, knowledge.topKeywords);
  const accuracy = scoreAccuracy(submissionText, knowledge.concepts);
  const reasoning = scoreReasoning(submissionText);
  const structure = scoreStructure(submissionText);
  const weightedScore = Math.round(
    coverage.score * weights.coverage +
      accuracy * weights.accuracy +
      reasoning * weights.reasoning +
      structure * weights.structure
  );
  const missingTerms = knowledge.topKeywords
    .slice(0, 8)
    .map((entry) => entry.term)
    .filter((term) => !coverage.present.includes(term));
  const strengths = [
    coverage.present.length
      ? `Uses source vocabulary such as ${coverage.present
          .slice(0, 4)
          .map(titleCase)
          .join(", ")}.`
      : "The submission includes some content, but it needs more source vocabulary.",
    accuracy >= 70
      ? "The response aligns well with the main concepts from the imported material."
      : "Some ideas are present, but several source concepts are still underdeveloped.",
    reasoning >= 65
      ? "The explanation shows a visible reasoning chain instead of disconnected facts."
      : "The explanation would be stronger with clearer causal links and step words.",
  ];
  const gaps = [
    missingTerms.length
      ? `Add missing high-value terms such as ${missingTerms.slice(0, 4).map(titleCase).join(", ")}.`
      : "Most of the high-value vocabulary from the source is present.",
    accuracy < 70
      ? "Recheck the source and tighten factual links so the answer matches the imported material more closely."
      : "Keep checking for precise wording so strong ideas stay accurate.",
    structure < 70
      ? "Break the response into clearer steps, numbered answers, or short paragraphs."
      : "The structure is usable, but it can still be tightened for speed and clarity.",
  ];
  const confidence = Math.min(
    96,
    Math.max(58, Math.round(55 + countWords(submissionText) / 7 + knowledge.stats.wordCount / 180))
  );

  return {
    breakdown: [
      { label: "Coverage", value: coverage.score },
      { label: "Accuracy", value: accuracy },
      { label: "Reasoning", value: reasoning },
      { label: "Structure", value: structure },
    ],
    confidence,
    fileName: `${slugify(knowledge.title)}-grading-feedback.md`,
    gaps,
    letter: letterGradeFor(weightedScore),
    markdown: [
      `# Grading Feedback`,
      `## Score`,
      `${weightedScore}% (${letterGradeFor(weightedScore)})`,
      `## Breakdown`,
      `- Coverage: ${coverage.score}%`,
      `- Accuracy: ${accuracy}%`,
      `- Reasoning: ${reasoning}%`,
      `- Structure: ${structure}%`,
      `## Strengths`,
      ...strengths.map((item) => `- ${item}`),
      `## Needs Work`,
      ...gaps.map((item) => `- ${item}`),
      `## Summary`,
      `This local grade estimates how closely the submission matches the imported source material. It is strongest when the response uses the same core vocabulary, explains connections clearly, and covers the major concepts from the study pack.`,
    ].join("\n"),
    score: weightedScore,
    strengths,
    summary: `This submission earned ${weightedScore}% (${letterGradeFor(
      weightedScore
    )}) based on how well it covered the imported source ideas, matched the main concepts, showed reasoning, and stayed organized. Use the missing-term list and concept gaps as the next revision targets.`,
  };
};

const handleGradeSubmission = async () => {
  const sourceResult = await resolveKnowledgeFromInputs({
    emptyMessage:
      "Add grading reference files/text or build a Workspace study pack before grading.",
    files: Array.from(app.gradingSourceFiles.files || []),
    statusElement: app.gradingStatus,
    text: app.gradingSourceText.value,
    textName: "Grading source",
  });

  if (!sourceResult) {
    return;
  }

  const submissionParts = [];
  const file = app.gradeFile.files && app.gradeFile.files[0] ? app.gradeFile.files[0] : null;
  const pasted = cleanText(app.gradingText.value);

  if (!file && !pasted) {
    setStatus(app.gradingStatus, "Upload a file or paste a submission before grading.", "error");
    return;
  }

  if (file) {
    try {
      setStatus(app.gradingStatus, `Scanning ${file.name}...`, "working");
      submissionParts.push(await extractTextFromFile(file));
    } catch (error) {
      console.error(error);
      setStatus(
        app.gradingStatus,
        `Could not fully read ${file.name}. ${error.message || "Please try a clearer file."}`,
        "error"
      );
      return;
    }
  }

  if (pasted) {
    submissionParts.push(pasted);
  }

  const submissionText = cleanText(submissionParts.join("\n\n"));

  if (!submissionText) {
    setStatus(app.gradingStatus, "No readable submission text was found.", "error");
    return;
  }

  state.currentGrade = buildGradeReport(
    submissionText,
    app.gradingRubric.value,
    sourceResult.knowledge
  );
  renderGrade();
  const sourceLabel =
    sourceResult.origin === "direct" ? "tab-specific files/text" : "the Workspace study pack";
  setStatus(
    app.gradingStatus,
    `Submission graded at ${state.currentGrade.score}% (${state.currentGrade.letter}) using ${sourceLabel}.`,
    "ready"
  );
};

const handleDownloadAll = () => {
  if (!state.artifacts.length) {
    return;
  }

  const bundle = state.artifacts
    .slice()
    .reverse()
    .map((artifact) => `# ${artifact.title}\n\n${artifact.markdown}`)
    .join("\n\n---\n\n");

  downloadText(
    `${slugify((state.knowledge && state.knowledge.title) || "study-pack")}-bundle.md`,
    bundle
  );
};

const handlePrintCurrent = () => {
  const artifact = getCurrentArtifact();

  if (!artifact) {
    return;
  }

  const popup = window.open("", "_blank", "width=1000,height=900");

  if (!popup) {
    return;
  }

  popup.document.write(`
    <!DOCTYPE html>
    <html lang="en">
      <head>
        <meta charset="UTF-8" />
        <title>${escapeHtml(artifact.title)}</title>
        <style>
          body {
            margin: 40px auto;
            width: min(760px, calc(100% - 48px));
            font-family: Georgia, "Times New Roman", serif;
            color: #111;
            line-height: 1.7;
          }
          h1, h2, h3 { line-height: 1.12; }
          blockquote { padding-left: 14px; border-left: 3px solid #999; color: #444; }
          ul, ol { padding-left: 22px; }
        </style>
      </head>
      <body>${markdownToHtml(artifact.markdown)}</body>
    </html>
  `);
  popup.document.close();
  popup.focus();
  popup.print();
};

const bindTabEvents = () => {
  app.sideTabs.forEach((button, index) => {
    button.addEventListener("click", () => {
      setActiveTab(button.dataset.appTab, { focusTab: false });
    });

    button.addEventListener("keydown", (event) => {
      const lastIndex = app.sideTabs.length - 1;
      let targetIndex = index;

      if (event.key === "ArrowDown" || event.key === "ArrowRight") {
        targetIndex = index === lastIndex ? 0 : index + 1;
      } else if (event.key === "ArrowUp" || event.key === "ArrowLeft") {
        targetIndex = index === 0 ? lastIndex : index - 1;
      } else if (event.key === "Home") {
        targetIndex = 0;
      } else if (event.key === "End") {
        targetIndex = lastIndex;
      } else {
        return;
      }

      event.preventDefault();
      setActiveTab(app.sideTabs[targetIndex].dataset.appTab, { focusTab: true });
    });
  });

  app.jumpTabs.forEach((button) => {
    button.addEventListener("click", () => {
      setActiveTab(button.dataset.jumpTab, { focusTab: true });
    });
  });
};

const initializeLibraries = () => {
  if (window.pdfjsLib) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  }

  if (window.mermaid) {
    window.mermaid.initialize({
      startOnLoad: false,
      securityLevel: "loose",
      theme: "base",
      themeVariables: {
        fontFamily: "Space Grotesk, sans-serif",
        primaryColor: "#ffd17a",
        primaryBorderColor: "#2ca8db",
        lineColor: "#5fd0ff",
        secondaryColor: "#10243a",
        tertiaryColor: "#07111d",
      },
    });
  }
};

const bindEvents = () => {
  bindTabEvents();
  app.ingestTrigger.addEventListener("click", () => handleIngest());
  app.sampleTrigger.addEventListener("click", () => handleIngest({ useSample: true }));
  app.clearTrigger.addEventListener("click", resetWorkspace);
  app.generateTrigger.addEventListener("click", () => handleGenerateArtifact());
  app.answerTrigger.addEventListener("click", handleAnswerQuestion);
  app.refreshQuestions.addEventListener("click", handleRefreshQuestions);
  app.gradeTrigger.addEventListener("click", handleGradeSubmission);
  app.downloadCurrent.addEventListener("click", () => {
    const artifact = getCurrentArtifact();

    if (artifact) {
      downloadText(artifact.fileName, artifact.markdown);
    }
  });
  app.downloadAll.addEventListener("click", handleDownloadAll);
  app.printCurrent.addEventListener("click", handlePrintCurrent);
  app.downloadAnswer.addEventListener("click", () => {
    if (state.currentAnswer) {
      downloadText(`${slugify(state.currentAnswer.question)}-answer.md`, state.currentAnswer.markdown);
    }
  });
  app.downloadGrade.addEventListener("click", () => {
    if (state.currentGrade) {
      downloadText(state.currentGrade.fileName, state.currentGrade.markdown);
    }
  });
};

const initialize = () => {
  initializeLibraries();
  setActiveTab(state.activeTab);
  bindEvents();
  renderSourceList();
  renderSourceOverview();
  renderQuestionList();
  renderArtifactList();
  renderAnswer();
  renderGrade();
};

initialize();
