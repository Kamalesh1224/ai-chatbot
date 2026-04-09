const chatbotConfig = {
  companyName: "Northstar Labs",
  supportEmail: "hello@northstarlabs.co",
  phone: "+1 (415) 555-0138",
  defaultSuggestions: [
    "What services do you offer?",
    "How does onboarding work?",
    "Do you support PDF knowledge bases?",
    "Can I talk to a human?"
  ],
  knowledgeManifest: "knowledge-base/manifest.json",
  storageKeys: {
    conversation: "northstarConversation",
    currentTopic: "northstarCurrentTopic",
    lastResolvedTopic: "northstarLastResolvedTopic",
    activeTopic: "northstarActiveTopic",
    adminDocuments: "northstarAdminDocuments",
    leads: "northstarLeads"
  }
};

const STOP_WORDS = new Set([
  "a", "an", "and", "are", "as", "at", "be", "been", "being", "but", "by",
  "can", "do", "does", "for", "from", "had", "has", "have", "how", "if", "in",
  "into", "is", "it", "its", "just", "more", "of", "on", "or", "our", "so",
  "that", "the", "their", "them", "there", "these", "they", "this", "to",
  "too", "was", "we", "were", "what", "when", "where", "which", "who", "why",
  "will", "with", "you", "your"
]);

const TOPIC_DEFINITIONS = [
  {
    id: "all",
    label: "All topics",
    shortLabel: "All",
    keywords: [],
    suggestions: [
      "What services do you offer?",
      "How does the product handle knowledge retrieval?",
      "What are your support hours?",
      "What is the PTO policy?"
    ]
  },
  {
    id: "general",
    label: "General",
    shortLabel: "General",
    keywords: [
      "company",
      "pricing",
      "discovery",
      "scope",
      "demo",
      "proposal",
      "quote",
      "commercial",
      "timeline",
      "engagement"
    ],
    suggestions: [
      "What services do you offer?",
      "How long do projects take?",
      "Can I book a demo?"
    ]
  },
  {
    id: "product",
    label: "Product",
    shortLabel: "Product",
    keywords: [
      "product",
      "feature",
      "chatbot",
      "widget",
      "knowledge",
      "retrieval",
      "source",
      "upload",
      "admin",
      "memory",
      "pdf",
      "quick reply",
      "document"
    ],
    suggestions: [
      "How does the product handle knowledge retrieval?",
      "Do you support PDF knowledge bases?",
      "Can admins upload new docs?"
    ]
  },
  {
    id: "support",
    label: "Support",
    shortLabel: "Support",
    keywords: [
      "support",
      "help",
      "issue",
      "urgent",
      "response time",
      "ticket",
      "incident",
      "hours",
      "bug",
      "broken",
      "escalation"
    ],
    suggestions: [
      "What are your support hours?",
      "How fast are response times?",
      "Can I talk to a human?"
    ]
  },
  {
    id: "hr",
    label: "HR",
    shortLabel: "HR",
    keywords: [
      "hr",
      "human resources",
      "employee",
      "vacation",
      "pto",
      "leave",
      "sick",
      "benefits",
      "hiring",
      "recruiting",
      "interview",
      "manager",
      "policy"
    ],
    suggestions: [
      "What is the PTO policy?",
      "How does employee onboarding work?",
      "Who handles recruiting questions?"
    ]
  }
];

const TOPICS_BY_ID = Object.fromEntries(
  TOPIC_DEFINITIONS.map((topicDefinition) => [topicDefinition.id, topicDefinition])
);

const DEFAULT_TOPIC_ID = "general";

const messagesEl = document.getElementById("chatbot-messages");
const suggestionsEl = document.getElementById("chatbot-suggestions");
const suggestionsBlockEl = document.getElementById("chatbot-suggestions-block");
const chatForm = document.getElementById("chat-form");
const chatInput = document.getElementById("chat-input");
const chatSubmitButton = chatForm.querySelector('button[type="submit"]');
const leadForm = document.getElementById("lead-form");
const leadStatus = document.getElementById("lead-status");
const leadMessage = document.getElementById("lead-message");
const chatbotPanel = document.getElementById("chatbot-panel");
const chatToggle = document.getElementById("chat-toggle");
const chatClose = document.getElementById("chatbot-close");
const openChatbotButton = document.getElementById("open-chatbot");
const knowledgeStatusEl = document.getElementById("knowledge-status");
const knowledgeDocsEl = document.getElementById("knowledge-docs");
const memoryStripEl = document.getElementById("memory-strip");
const topicFiltersEl = document.getElementById("topic-filters");
const knowledgeUploadEl = document.getElementById("knowledge-upload");
const uploadKnowledgeTriggerEl = document.getElementById("upload-knowledge-trigger");
const clearMemoryButtonEl = document.getElementById("clear-memory");
const adminTopicSelectEl = document.getElementById("admin-topic-select");
const adminUploadTriggerEl = document.getElementById("admin-upload-trigger");
const adminResetDocsEl = document.getElementById("admin-reset-docs");
const adminUploadStatusEl = document.getElementById("admin-upload-status");
const adminKnowledgeUploadEl = document.getElementById("admin-knowledge-upload");
const adminDocListEl = document.getElementById("admin-doc-list");
const adminUploadedCountEl = document.getElementById("admin-uploaded-count");
const adminUploadedListEl = document.getElementById("admin-uploaded-list");
const adminSelectedDocEl = document.getElementById("admin-selected-doc");
const adminEditorFormEl = document.getElementById("admin-editor-form");
const adminEditorIdEl = document.getElementById("admin-editor-id");
const adminEditorTitleEl = document.getElementById("admin-editor-title");
const adminEditorTopicEl = document.getElementById("admin-editor-topic");
const adminEditorContentEl = document.getElementById("admin-editor-content");
const adminEditorStatusEl = document.getElementById("admin-editor-status");
const adminCancelEditEl = document.getElementById("admin-cancel-edit");
const adminEditorSaveButtonEl = adminEditorFormEl?.querySelector('button[type="submit"]');
const topicSummaryEl = document.getElementById("topic-summary");

const state = {
  unresolvedCount: 0,
  isReplyPending: false,
  typingIndicatorEl: null,
  activeSuggestions: [...chatbotConfig.defaultSuggestions],
  documents: [],
  chunks: [],
  inverseDocumentFrequency: {},
  activeTopic: loadStoredJson(chatbotConfig.storageKeys.activeTopic, "all"),
  lastResolvedTopic: loadStoredJson(chatbotConfig.storageKeys.lastResolvedTopic, ""),
  currentTopic: loadStoredJson(chatbotConfig.storageKeys.currentTopic, ""),
  conversation: loadStoredJson(chatbotConfig.storageKeys.conversation, []),
  selectedAdminDocumentId: ""
};

function loadStoredJson(key, fallbackValue) {
  try {
    const stored = localStorage.getItem(key);
    return stored ? JSON.parse(stored) : fallbackValue;
  } catch (error) {
    return fallbackValue;
  }
}

function saveStoredJson(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch (error) {
    // Ignore storage failures so chat keeps working.
  }
}

function sanitizeTopicId(topicId, { allowAll = true } = {}) {
  const normalizedTopicId = String(topicId || "")
    .toLowerCase()
    .trim()
    .replace(/[^a-z]/g, "");

  if (allowAll && normalizedTopicId === "all") {
    return "all";
  }

  if (TOPICS_BY_ID[normalizedTopicId] && normalizedTopicId !== "all") {
    return normalizedTopicId;
  }

  return DEFAULT_TOPIC_ID;
}

function getTopicDefinition(topicId) {
  return TOPICS_BY_ID[sanitizeTopicId(topicId)] || TOPICS_BY_ID[DEFAULT_TOPIC_ID];
}

function getTopicLabel(topicId) {
  return getTopicDefinition(topicId).label;
}

function getTopicSuggestions(topicId) {
  const normalizedTopicId = sanitizeTopicId(topicId);
  if (normalizedTopicId === "all") {
    return state.activeSuggestions;
  }

  return getTopicDefinition(normalizedTopicId).suggestions || state.activeSuggestions;
}

function getDefaultQuickReplies() {
  return state.activeTopic !== "all"
    ? getTopicSuggestions(state.activeTopic)
    : state.activeSuggestions;
}

function createDocumentId(prefix = "doc") {
  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
}

function prepareDocumentEntry(documentEntry, overrides = {}) {
  const mergedEntry = {
    ...documentEntry,
    ...overrides
  };

  return {
    id: mergedEntry.id || createDocumentId("doc"),
    title: String(mergedEntry.title || "Untitled document").trim(),
    type: String(mergedEntry.type || "text").toLowerCase(),
    origin: String(mergedEntry.origin || "manifest").toLowerCase(),
    topic: sanitizeTopicId(mergedEntry.topic, { allowAll: false }),
    text: String(mergedEntry.text || "").trim(),
    createdAt: mergedEntry.createdAt || new Date().toISOString(),
    updatedAt: mergedEntry.updatedAt || mergedEntry.createdAt || new Date().toISOString()
  };
}

function getStoredAdminDocuments() {
  const storedDocuments = loadStoredJson(chatbotConfig.storageKeys.adminDocuments, []);
  return Array.isArray(storedDocuments)
    ? storedDocuments
        .map((documentEntry) => prepareDocumentEntry(documentEntry, { origin: "admin" }))
        .filter((documentEntry) => documentEntry.text)
    : [];
}

function saveAdminDocuments() {
  const adminDocuments = state.documents.filter((documentEntry) => documentEntry.origin === "admin");
  saveStoredJson(chatbotConfig.storageKeys.adminDocuments, adminDocuments);
}

function countDocumentsByTopic() {
  return state.documents.reduce((counts, documentEntry) => {
    const topicId = sanitizeTopicId(documentEntry.topic, { allowAll: false });
    counts[topicId] = (counts[topicId] || 0) + 1;
    return counts;
  }, {});
}

function countChunksByTopic() {
  return state.chunks.reduce((counts, chunk) => {
    counts[chunk.topic] = (counts[chunk.topic] || 0) + 1;
    return counts;
  }, {});
}

function normalize(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function tokenize(text) {
  return normalize(text)
    .split(" ")
    .filter((token) => token.length > 2 && !STOP_WORDS.has(token));
}

function countTokens(tokens) {
  const counts = {};
  tokens.forEach((token) => {
    counts[token] = (counts[token] || 0) + 1;
  });
  return counts;
}

function unique(items) {
  return Array.from(new Set(items.filter(Boolean)));
}

function wait(duration) {
  return new Promise((resolve) => {
    window.setTimeout(resolve, duration);
  });
}

function truncate(text, maxLength = 170) {
  if (!text) {
    return "";
  }

  return text.length <= maxLength ? text : `${text.slice(0, maxLength - 1).trim()}...`;
}

function humanizeFileName(fileName) {
  return fileName
    .replace(/\.[^.]+$/, "")
    .replace(/[-_]+/g, " ")
    .replace(/\s+/g, " ")
    .replace(/\b\w/g, (match) => match.toUpperCase())
    .trim();
}

function updateKnowledgeStatus(message) {
  const isLoading = /\b(loading|importing|uploading|indexing|preparing)\b/i.test(message);
  knowledgeStatusEl.textContent = message;
  knowledgeStatusEl.classList.toggle("is-loading", isLoading);
  knowledgeStatusEl.setAttribute("aria-busy", String(isLoading));
}

function renderKnowledgeDocuments() {
  knowledgeDocsEl.innerHTML = "";

  if (!state.documents.length) {
    const emptyChip = document.createElement("span");
    emptyChip.className = "doc-chip doc-chip-empty";
    emptyChip.textContent = "No documents loaded yet";
    knowledgeDocsEl.appendChild(emptyChip);
    return;
  }

  state.documents.forEach((documentEntry) => {
    const chip = document.createElement("span");
    chip.className = "doc-chip";
    chip.textContent = `${documentEntry.title} · ${documentEntry.type.toUpperCase()}`;
    knowledgeDocsEl.appendChild(chip);
  });
}

function updateMemoryStrip() {
  const userTurns = state.conversation.filter((turn) => turn.role === "user");
  if (!userTurns.length) {
    memoryStripEl.textContent = "Memory is empty. Follow-up questions will start using context after the first question.";
    return;
  }

  const latestQuestion = userTurns[userTurns.length - 1].text;
  const topicText = state.currentTopic ? `Topic: ${state.currentTopic}. ` : "";
  const turnLabel = userTurns.length === 1 ? "question" : "questions";
  memoryStripEl.textContent = `${topicText}Remembering ${userTurns.length} recent ${turnLabel}. Latest: ${truncate(latestQuestion, 88)}`;
}

function scrollMessagesToBottom() {
  messagesEl.scrollTop = messagesEl.scrollHeight;
}

function addMessage(sender, payload) {
  const messageData = typeof payload === "string"
    ? { paragraphs: [payload] }
    : Array.isArray(payload)
      ? { paragraphs: payload }
      : payload;

  const wrapper = document.createElement("div");
  wrapper.className = `message ${sender}`;

  (messageData.paragraphs || []).forEach((paragraph) => {
    const text = document.createElement("p");
    text.textContent = paragraph;
    wrapper.appendChild(text);
  });

  if (sender === "bot" && Array.isArray(messageData.sources) && messageData.sources.length) {
    const sourcesContainer = document.createElement("div");
    sourcesContainer.className = "message-sources";

    messageData.sources.forEach((source) => {
      const sourceCard = document.createElement("div");
      sourceCard.className = "source-card";

      const sourceTitle = document.createElement("span");
      sourceTitle.className = "source-title";
      sourceTitle.textContent = source.title;

      const sourceExcerpt = document.createElement("span");
      sourceExcerpt.className = "source-excerpt";
      sourceExcerpt.textContent = truncate(source.excerpt, 150);

      sourceCard.appendChild(sourceTitle);
      sourceCard.appendChild(sourceExcerpt);
      sourcesContainer.appendChild(sourceCard);
    });

    wrapper.appendChild(sourcesContainer);
  }

  messagesEl.appendChild(wrapper);
  scrollMessagesToBottom();
}

function setSuggestions(items) {
  suggestionsEl.innerHTML = "";
  const quickReplies = unique(items).slice(0, 4);

  suggestionsBlockEl.hidden = !quickReplies.length;
  if (!quickReplies.length) {
    return;
  }

  quickReplies.forEach((item) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "suggestion-chip";
    button.textContent = item;
    button.disabled = state.isReplyPending;
    button.addEventListener("click", () => {
      handleUserMessage(item);
    });
    suggestionsEl.appendChild(button);
  });
}

function setChatBusy(isBusy) {
  state.isReplyPending = isBusy;
  chatInput.disabled = isBusy;
  chatSubmitButton.disabled = isBusy;
  uploadKnowledgeTriggerEl.disabled = isBusy;
  clearMemoryButtonEl.disabled = isBusy;

  suggestionsEl.querySelectorAll("button").forEach((button) => {
    button.disabled = isBusy;
  });
}

function showTypingIndicator() {
  removeTypingIndicator();

  const wrapper = document.createElement("div");
  wrapper.className = "message bot typing";
  wrapper.setAttribute("role", "status");
  wrapper.setAttribute("aria-label", `${chatbotConfig.companyName} assistant is typing`);

  const avatar = document.createElement("span");
  avatar.className = "typing-avatar";
  avatar.textContent = chatbotConfig.companyName.charAt(0);

  const content = document.createElement("div");
  content.className = "typing-content";

  const label = document.createElement("span");
  label.className = "typing-label";
  label.textContent = `${chatbotConfig.companyName} is typing`;

  const indicator = document.createElement("div");
  indicator.className = "typing-indicator";

  for (let index = 0; index < 3; index += 1) {
    const dot = document.createElement("span");
    dot.className = "typing-dot";
    dot.style.animationDelay = `${index * 0.16}s`;
    indicator.appendChild(dot);
  }

  content.appendChild(label);
  content.appendChild(indicator);
  wrapper.appendChild(avatar);
  wrapper.appendChild(content);
  messagesEl.appendChild(wrapper);
  state.typingIndicatorEl = wrapper;
  scrollMessagesToBottom();
}

function removeTypingIndicator() {
  if (state.typingIndicatorEl) {
    state.typingIndicatorEl.remove();
    state.typingIndicatorEl = null;
  }
}

function getTypingDelay(reply) {
  const paragraphLength = (reply.paragraphs || []).join(" ").length;
  return Math.min(1100, Math.max(420, 220 + paragraphLength * 4));
}

function renderKnowledgeDocuments() {
  knowledgeDocsEl.innerHTML = "";

  if (!state.documents.length) {
    const emptyChip = document.createElement("span");
    emptyChip.className = "doc-chip doc-chip-empty";
    emptyChip.textContent = "No documents loaded yet";
    knowledgeDocsEl.appendChild(emptyChip);
    return;
  }

  state.documents.forEach((documentEntry) => {
    const chip = document.createElement("span");
    chip.className = "doc-chip";

    const originLabel = documentEntry.origin === "admin"
      ? "Admin"
      : documentEntry.origin === "session"
        ? "Session"
        : "Starter";

    chip.textContent = `${documentEntry.title} | ${getTopicLabel(documentEntry.topic)} | ${originLabel}`;
    knowledgeDocsEl.appendChild(chip);
  });
}

function updateMemoryStrip() {
  const focusText = `Focus: ${getTopicLabel(state.activeTopic)}.`;
  const routedTopicText = state.lastResolvedTopic
    ? ` Last answer topic: ${getTopicLabel(state.lastResolvedTopic)}.`
    : "";
  const userTurns = state.conversation.filter((turn) => turn.role === "user");

  if (!userTurns.length) {
    memoryStripEl.textContent = `${focusText}${routedTopicText} Memory is empty. Follow-up questions will start using context after the first question.`;
    return;
  }

  const latestQuestion = userTurns[userTurns.length - 1].text;
  const sourceText = state.currentTopic ? ` Latest source: ${state.currentTopic}.` : "";
  const turnLabel = userTurns.length === 1 ? "question" : "questions";
  memoryStripEl.textContent = `${focusText}${routedTopicText} Remembering ${userTurns.length} recent ${turnLabel}. Latest: ${truncate(latestQuestion, 88)}.${sourceText}`;
}

function renderTopicFilters() {
  topicFiltersEl.innerHTML = "";

  TOPIC_DEFINITIONS.forEach((topicDefinition) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `topic-filter${state.activeTopic === topicDefinition.id ? " is-active" : ""}`;
    button.textContent = topicDefinition.shortLabel;
    button.disabled = state.isReplyPending;
    button.setAttribute("aria-pressed", String(state.activeTopic === topicDefinition.id));
    button.addEventListener("click", () => {
      setActiveTopic(topicDefinition.id, { announce: true });
    });
    topicFiltersEl.appendChild(button);
  });
}

function setActiveTopic(topicId, { announce = false } = {}) {
  state.activeTopic = sanitizeTopicId(topicId);
  state.unresolvedCount = 0;
  saveStoredJson(chatbotConfig.storageKeys.activeTopic, state.activeTopic);
  hideLeadForm();
  renderTopicFilters();
  updateMemoryStrip();
  setSuggestions(getDefaultQuickReplies());

  if (announce) {
    addMessage("bot", [
      state.activeTopic === "all"
        ? "Chat focus is back on all topics."
        : `Chat focus is now set to ${getTopicLabel(state.activeTopic)}.`
    ]);
  }
}

function renderTopicSummary() {
  topicSummaryEl.innerHTML = "";
  const documentCounts = countDocumentsByTopic();
  const chunkCounts = countChunksByTopic();

  TOPIC_DEFINITIONS
    .filter((topicDefinition) => topicDefinition.id !== "all")
    .forEach((topicDefinition) => {
      const summaryCard = document.createElement("article");
      summaryCard.className = "topic-summary-card";

      const label = document.createElement("p");
      label.className = "topic-summary-label";
      label.textContent = topicDefinition.label;

      const totalDocs = document.createElement("strong");
      totalDocs.className = "topic-summary-count";
      totalDocs.textContent = `${documentCounts[topicDefinition.id] || 0} docs`;

      const chunkMeta = document.createElement("span");
      chunkMeta.className = "topic-summary-meta";
      chunkMeta.textContent = `${chunkCounts[topicDefinition.id] || 0} indexed sections`;

      summaryCard.appendChild(label);
      summaryCard.appendChild(totalDocs);
      summaryCard.appendChild(chunkMeta);
      topicSummaryEl.appendChild(summaryCard);
    });
}

function getAdminDocuments() {
  return state.documents
    .filter((documentEntry) => documentEntry.origin === "admin")
    .sort((left, right) => {
      const leftTimestamp = new Date(left.updatedAt || left.createdAt || 0).getTime();
      const rightTimestamp = new Date(right.updatedAt || right.createdAt || 0).getTime();
      return rightTimestamp - leftTimestamp;
    });
}

function getSelectedAdminDocument() {
  return state.documents.find((documentEntry) => {
    return documentEntry.origin === "admin" && documentEntry.id === state.selectedAdminDocumentId;
  }) || null;
}

function formatAdminTimestamp(timestamp) {
  const parsedTimestamp = new Date(timestamp || "");
  if (Number.isNaN(parsedTimestamp.getTime())) {
    return "Saved recently";
  }

  return `Updated ${parsedTimestamp.toLocaleString([], {
    month: "short",
    day: "numeric",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit"
  })}`;
}

function setStatusMessage(element, message) {
  const isLoading = /\b(loading|importing|uploading|indexing|preparing|saving|updating|deleting|removing)\b/i.test(message);
  element.textContent = message;
  element.classList.toggle("is-loading", isLoading);
  element.setAttribute("aria-busy", String(isLoading));
}

function setAdminEditorStatus(message) {
  setStatusMessage(adminEditorStatusEl, message);
}

function renderAdminControlPanel() {
  const adminDocuments = getAdminDocuments();
  const selectedDocument = getSelectedAdminDocument();
  const activeDocument = selectedDocument || null;

  if (!activeDocument && state.selectedAdminDocumentId) {
    state.selectedAdminDocumentId = "";
  }

  adminUploadedCountEl.textContent = adminDocuments.length
    ? `${adminDocuments.length} admin upload${adminDocuments.length === 1 ? "" : "s"} available in this browser.`
    : "No admin uploads yet.";

  adminUploadedListEl.innerHTML = "";

  if (!adminDocuments.length) {
    const emptyState = document.createElement("div");
    emptyState.className = "admin-uploaded-card admin-doc-empty";
    emptyState.textContent = "Upload documents to manage them here.";
    adminUploadedListEl.appendChild(emptyState);
  } else {
    adminDocuments.forEach((documentEntry) => {
      const card = document.createElement("article");
      card.className = `admin-uploaded-card${documentEntry.id === state.selectedAdminDocumentId ? " is-selected" : ""}`;

      const header = document.createElement("div");
      header.className = "admin-uploaded-header";

      const info = document.createElement("div");
      info.className = "admin-doc-title-group";

      const title = document.createElement("h4");
      title.className = "admin-doc-title";
      title.textContent = documentEntry.title;

      const meta = document.createElement("div");
      meta.className = "admin-doc-meta";

      [
        { text: getTopicLabel(documentEntry.topic), extraClass: "" },
        { text: documentEntry.type.toUpperCase(), extraClass: "admin-doc-badge-muted" }
      ].forEach((badgeEntry) => {
        const badge = document.createElement("span");
        badge.className = `admin-doc-badge ${badgeEntry.extraClass}`.trim();
        badge.textContent = badgeEntry.text;
        meta.appendChild(badge);
      });

      const timestamp = document.createElement("p");
      timestamp.className = "admin-uploaded-time";
      timestamp.textContent = formatAdminTimestamp(documentEntry.updatedAt || documentEntry.createdAt);

      info.appendChild(title);
      info.appendChild(meta);
      info.appendChild(timestamp);

      const actions = document.createElement("div");
      actions.className = "admin-doc-actions";

      const editButton = document.createElement("button");
      editButton.type = "button";
      editButton.className = "admin-doc-action";
      editButton.textContent = "Edit";
      editButton.dataset.adminAction = "edit";
      editButton.dataset.documentId = documentEntry.id;
      editButton.disabled = state.isReplyPending;

      const deleteButton = document.createElement("button");
      deleteButton.type = "button";
      deleteButton.className = "admin-doc-action admin-doc-action-danger";
      deleteButton.textContent = "Delete";
      deleteButton.dataset.adminAction = "delete";
      deleteButton.dataset.documentId = documentEntry.id;
      deleteButton.disabled = state.isReplyPending;

      actions.appendChild(editButton);
      actions.appendChild(deleteButton);

      header.appendChild(info);
      header.appendChild(actions);

      const excerpt = document.createElement("p");
      excerpt.className = "admin-doc-excerpt";
      excerpt.textContent = truncate(documentEntry.text, 200);

      card.appendChild(header);
      card.appendChild(excerpt);
      adminUploadedListEl.appendChild(card);
    });
  }

  const hasSelectedDocument = Boolean(activeDocument);
  adminSelectedDocEl.textContent = hasSelectedDocument ? activeDocument.title : "No document selected";
  adminEditorFormEl.classList.toggle("is-empty", !hasSelectedDocument);
  adminEditorIdEl.value = activeDocument?.id || "";
  adminEditorTitleEl.value = activeDocument?.title || "";
  adminEditorTopicEl.value = activeDocument?.topic || sanitizeTopicId(adminTopicSelectEl.value, { allowAll: false });
  adminEditorContentEl.value = activeDocument?.text || "";

  [
    adminEditorTitleEl,
    adminEditorTopicEl,
    adminEditorContentEl,
    adminEditorSaveButtonEl
  ].forEach((field) => {
    field.disabled = state.isReplyPending || !hasSelectedDocument;
  });

  adminCancelEditEl.disabled = state.isReplyPending || !hasSelectedDocument;

  if (!hasSelectedDocument) {
    setAdminEditorStatus(
      adminDocuments.length
        ? "Select an uploaded document to update it."
        : "Upload a document to start managing the admin knowledge base."
    );
  }
}

function setSelectedAdminDocument(documentId) {
  const selectedDocument = state.documents.find((documentEntry) => {
    return documentEntry.origin === "admin" && documentEntry.id === documentId;
  }) || null;

  state.selectedAdminDocumentId = selectedDocument ? selectedDocument.id : "";
  renderAdminControlPanel();

  if (selectedDocument) {
    setAdminEditorStatus(`${selectedDocument.title} is ready to edit. ${formatAdminTimestamp(selectedDocument.updatedAt || selectedDocument.createdAt)}.`);
    if (!state.isReplyPending) {
      adminEditorTitleEl.focus();
      adminEditorTitleEl.select();
    }
  }
}

function updateAdminDocument(documentId, updates) {
  const documentIndex = state.documents.findIndex((documentEntry) => {
    return documentEntry.origin === "admin" && documentEntry.id === documentId;
  });

  if (documentIndex < 0) {
    return null;
  }

  const currentDocument = state.documents[documentIndex];
  const updatedDocument = prepareDocumentEntry(currentDocument, {
    title: String(updates.title || currentDocument.title).trim(),
    topic: sanitizeTopicId(updates.topic || currentDocument.topic, { allowAll: false }),
    text: String(updates.text || currentDocument.text).trim(),
    updatedAt: new Date().toISOString()
  });

  state.documents[documentIndex] = updatedDocument;
  state.selectedAdminDocumentId = updatedDocument.id;
  rebuildKnowledgeIndex();
  saveAdminDocuments();
  setAdminStatus(`Updated ${updatedDocument.title} in the admin knowledge library.`);
  setAdminEditorStatus(`Saved changes to ${updatedDocument.title}.`);
  return updatedDocument;
}

function renderAdminDocumentList() {
  adminDocListEl.innerHTML = "";
  const hasAdminDocuments = state.documents.some((documentEntry) => documentEntry.origin === "admin");
  adminResetDocsEl.disabled = state.isReplyPending || !hasAdminDocuments;

  const libraryDocuments = state.documents
    .filter((documentEntry) => documentEntry.origin !== "session")
    .sort((left, right) => left.title.localeCompare(right.title));

  if (!libraryDocuments.length) {
    const emptyState = document.createElement("div");
    emptyState.className = "admin-doc-card admin-doc-empty";
    emptyState.textContent = "No knowledge documents are available yet.";
    adminDocListEl.appendChild(emptyState);
    return;
  }

  libraryDocuments.forEach((documentEntry) => {
    const card = document.createElement("article");
    card.className = "admin-doc-card";

    const header = document.createElement("div");
    header.className = "admin-doc-header";

    const titleGroup = document.createElement("div");
    titleGroup.className = "admin-doc-title-group";

    const title = document.createElement("h4");
    title.className = "admin-doc-title";
    title.textContent = documentEntry.title;

    const meta = document.createElement("div");
    meta.className = "admin-doc-meta";

    [
      { text: getTopicLabel(documentEntry.topic), extraClass: "" },
      { text: documentEntry.type.toUpperCase(), extraClass: "admin-doc-badge-muted" },
      {
        text: documentEntry.origin === "admin" ? "Admin upload" : "Starter",
        extraClass: documentEntry.origin === "admin" ? "admin-doc-badge-strong" : "admin-doc-badge-muted"
      }
    ].forEach((badgeEntry) => {
      const badge = document.createElement("span");
      badge.className = `admin-doc-badge ${badgeEntry.extraClass}`.trim();
      badge.textContent = badgeEntry.text;
      meta.appendChild(badge);
    });

    titleGroup.appendChild(title);
    titleGroup.appendChild(meta);
    header.appendChild(titleGroup);

    const excerpt = document.createElement("p");
    excerpt.className = "admin-doc-excerpt";
    excerpt.textContent = truncate(documentEntry.text, 220);

    card.appendChild(header);
    card.appendChild(excerpt);
    adminDocListEl.appendChild(card);
  });
}

function setAdminStatus(message) {
  setStatusMessage(adminUploadStatusEl, message);
}

function setChatBusy(isBusy) {
  state.isReplyPending = isBusy;
  chatInput.disabled = isBusy;
  chatSubmitButton.disabled = isBusy;
  uploadKnowledgeTriggerEl.disabled = isBusy;
  clearMemoryButtonEl.disabled = isBusy;
  adminUploadTriggerEl.disabled = isBusy;
  adminResetDocsEl.disabled = isBusy || !state.documents.some((documentEntry) => documentEntry.origin === "admin");
  adminTopicSelectEl.disabled = isBusy;

  suggestionsEl.querySelectorAll("button").forEach((button) => {
    button.disabled = isBusy;
  });

  topicFiltersEl.querySelectorAll("button").forEach((button) => {
    button.disabled = isBusy;
  });

  adminUploadedListEl.querySelectorAll("button").forEach((button) => {
    button.disabled = isBusy;
  });

  adminDocListEl.querySelectorAll("button").forEach((button) => {
    button.disabled = isBusy;
  });

  [
    adminEditorTitleEl,
    adminEditorTopicEl,
    adminEditorContentEl,
    adminEditorSaveButtonEl,
    adminCancelEditEl
  ].forEach((field) => {
    field.disabled = isBusy || !state.selectedAdminDocumentId;
  });
}

function showLeadForm(prefillMessage = "") {
  leadForm.classList.remove("hidden");
  if (prefillMessage) {
    leadMessage.value = prefillMessage;
  }
}

function hideLeadForm() {
  leadForm.classList.add("hidden");
  leadStatus.textContent = "";
}

function openChat() {
  chatbotPanel.classList.remove("is-hidden");
  chatToggle.setAttribute("aria-expanded", "true");
  chatInput.focus();
}

function closeChat() {
  chatbotPanel.classList.add("is-hidden");
  chatToggle.setAttribute("aria-expanded", "false");
}

function persistConversation() {
  saveStoredJson(chatbotConfig.storageKeys.conversation, state.conversation);
  saveStoredJson(chatbotConfig.storageKeys.currentTopic, state.currentTopic);
}

function getOptionalNormalizedTopicId(topicId) {
  if (!topicId) {
    return "";
  }

  const normalizedTopicId = sanitizeTopicId(topicId);
  return normalizedTopicId === "all" ? "" : normalizedTopicId;
}

function rememberTurn(role, text, sources = [], metadata = {}) {
  const topicId = getOptionalNormalizedTopicId(metadata.topicId);
  state.conversation.push({
    role,
    text: truncate(text, role === "assistant" ? 420 : 320),
    sources: Array.isArray(sources)
      ? sources
          .map((source) => (typeof source === "string" ? source : source.title))
          .filter(Boolean)
          .slice(0, 4)
      : [],
    topicId,
    createdAt: new Date().toISOString()
  });

  state.conversation = state.conversation.slice(-24);
  persistConversation();
  updateMemoryStrip();
}

function clearConversationMemory() {
  state.conversation = [];
  state.currentTopic = "";
  state.unresolvedCount = 0;
  persistConversation();
  updateMemoryStrip();
}

function includesAny(text, keywords) {
  return keywords.some((keyword) => {
    if (keyword.includes(" ")) {
      return text.includes(keyword);
    }

    const escapedKeyword = keyword.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    return new RegExp(`\\b${escapedKeyword}\\b`).test(text);
  });
}

function hasContextReference(normalizedText) {
  return /\b(it|that|this|they|them|those|these|there|here|same|too|also|one|ones)\b/.test(normalizedText);
}

function isFollowUpQuestion(normalizedText) {
  const tokenCount = normalizedText.split(" ").filter(Boolean).length;
  return (
    tokenCount <= 8 ||
    /^(and|also|what about|how about|what else|tell me more|more on that|does it|does that|do they|can it|can they|can that|will it|is it|is that|are they|that|those|these|it|they|them|same for|what if|how much|how long|who handles)\b/.test(normalizedText) ||
    (tokenCount <= 14 && hasContextReference(normalizedText))
  );
}

function buildContextualQuery(message) {
  const normalizedMessage = normalize(message);
  const previousUserTurns = state.conversation
    .filter((turn) => turn.role === "user")
    .slice(-2)
    .map((turn) => turn.text);

  const useMemory = isFollowUpQuestion(normalizedMessage) && previousUserTurns.length > 0;
  const parts = [message];

  if (useMemory) {
    previousUserTurns.forEach((question) => {
      parts.push(question);
    });
    if (state.currentTopic) {
      parts.push(state.currentTopic);
    }
  }

  return {
    query: parts.join(" "),
    usedMemory: useMemory
  };
}

function splitIntoChunks(text) {
  const paragraphs = String(text || "")
    .replace(/\r/g, "")
    .split(/\n{2,}/)
    .map((paragraph) => paragraph.replace(/\n+/g, " ").replace(/\s+/g, " ").trim())
    .filter((paragraph) => paragraph.length > 30);

  if (!paragraphs.length) {
    return [];
  }

  const chunks = [];
  let currentChunk = "";

  paragraphs.forEach((paragraph) => {
    if (!currentChunk) {
      currentChunk = paragraph;
      return;
    }

    if (`${currentChunk} ${paragraph}`.length <= 520) {
      currentChunk = `${currentChunk} ${paragraph}`;
      return;
    }

    chunks.push(currentChunk);
    currentChunk = paragraph;
  });

  if (currentChunk) {
    chunks.push(currentChunk);
  }

  return chunks;
}

function rebuildKnowledgeIndex() {
  const chunks = [];

  state.documents.forEach((documentEntry) => {
    splitIntoChunks(documentEntry.text).forEach((chunkText, index) => {
      const tokens = tokenize(`${documentEntry.title} ${chunkText}`);
      if (!tokens.length) {
        return;
      }

      chunks.push({
        id: `${documentEntry.id}-${index + 1}`,
        documentId: documentEntry.id,
        documentTitle: documentEntry.title,
        documentType: documentEntry.type,
        text: chunkText,
        normalized: normalize(chunkText),
        tokenCounts: countTokens(tokens),
        uniqueTokens: unique(tokens)
      });
    });
  });

  const documentFrequency = {};
  chunks.forEach((chunk) => {
    chunk.uniqueTokens.forEach((token) => {
      documentFrequency[token] = (documentFrequency[token] || 0) + 1;
    });
  });

  const totalChunks = chunks.length || 1;
  state.inverseDocumentFrequency = Object.fromEntries(
    Object.entries(documentFrequency).map(([token, frequency]) => {
      return [token, Math.log((1 + totalChunks) / (1 + frequency)) + 1];
    })
  );
  state.chunks = chunks;

  renderKnowledgeDocuments();

  if (!state.documents.length) {
    updateKnowledgeStatus("No knowledge loaded");
    return;
  }

  updateKnowledgeStatus(`${state.documents.length} document${state.documents.length === 1 ? "" : "s"} ready`);
}

function mergeKnowledgeDocuments(newDocuments) {
  const seenKeys = new Set(state.documents.map((documentEntry) => `${documentEntry.title}::${documentEntry.text}`));

  newDocuments.forEach((documentEntry) => {
    const key = `${documentEntry.title}::${documentEntry.text}`;
    if (!seenKeys.has(key)) {
      state.documents.push(documentEntry);
      seenKeys.add(key);
    }
  });

  rebuildKnowledgeIndex();
}

function scoreChunk(chunk, queryTokens, normalizedQuery) {
  let score = 0;

  queryTokens.forEach((token) => {
    const tokenFrequency = chunk.tokenCounts[token] || 0;
    if (tokenFrequency) {
      score += tokenFrequency * (state.inverseDocumentFrequency[token] || 1);
    }
  });

  if (normalizedQuery && chunk.normalized.includes(normalizedQuery)) {
    score += 6;
  }

  return score;
}

function searchKnowledgeBase(query) {
  const normalizedQuery = normalize(query);
  const queryTokens = tokenize(query);

  if (!queryTokens.length || !state.chunks.length) {
    return [];
  }

  return state.chunks
    .map((chunk) => {
      return {
        chunk,
        score: scoreChunk(chunk, queryTokens, normalizedQuery)
      };
    })
    .filter((match) => match.score > 0)
    .sort((left, right) => right.score - left.score)
    .slice(0, 4);
}

function splitIntoSentences(text) {
  const matches = text.match(/[^.!?]+[.!?]?/g);
  return (matches || [text])
    .map((sentence) => sentence.trim())
    .filter(Boolean);
}

function buildSnippet(chunkText, queryTokens) {
  const scoredSentences = splitIntoSentences(chunkText)
    .map((sentence) => {
      const sentenceTokens = tokenize(sentence);
      const overlap = queryTokens.reduce((score, token) => {
        return score + (sentenceTokens.includes(token) ? 1 : 0);
      }, 0);

      return {
        sentence,
        overlap
      };
    })
    .sort((left, right) => right.overlap - left.overlap);

  if (scoredSentences.length && scoredSentences[0].overlap > 0) {
    return truncate(scoredSentences.slice(0, 2).map((entry) => entry.sentence).join(" "), 240);
  }

  return truncate(chunkText, 240);
}

function sourceSuggestions(documentTitle) {
  const normalizedTitle = normalize(documentTitle);

  if (normalizedTitle.includes("service")) {
    return [
      "Do you support PDF knowledge bases?",
      "How does onboarding work?",
      "What does the chatbot remember?"
    ];
  }

  if (normalizedTitle.includes("support")) {
    return [
      "What are your support hours?",
      "How fast are response times?",
      "Can I talk to a human?"
    ];
  }

  if (normalizedTitle.includes("commercial")) {
    return [
      "What does pricing usually look like?",
      "How long do projects take?",
      "Can I book a demo?"
    ];
  }

  return chatbotConfig.defaultSuggestions;
}

function buildDontKnowReply({ noKnowledgeLoaded = false, offerHumanFollowUp = false } = {}) {
  if (noKnowledgeLoaded) {
    return {
      paragraphs: [
        "I don't know yet because there isn't a knowledge base loaded.",
        "Use Load PDF / Text to add company material, or choose a quick reply if you want help with the next step."
      ],
      suggestions: [
        "How do I load documents?",
        "What can you do?",
        "Can I talk to a human?"
      ]
    };
  }

  if (offerHumanFollowUp) {
    return {
      paragraphs: [
        "I don't know based on the documents I have loaded right now.",
        "I've opened the follow-up form below so a teammate can help, or you can upload another document that covers this topic."
      ],
      suggestions: [
        "Can I talk to a human?",
        "How do I load documents?",
        "What services do you offer?"
      ],
      showLeadForm: true
    };
  }

  return {
    paragraphs: [
      "I don't know based on the documents I have loaded right now.",
      "Try a more specific phrase from the source material, or upload another PDF or text file that covers this topic."
    ],
    suggestions: [
      "How do I load documents?",
      "What services do you offer?",
      "Can I talk to a human?"
    ]
  };
}

function buildKnowledgeReply(message) {
  const normalizedMessage = normalize(message);

  if (!normalizedMessage) {
    state.unresolvedCount = 0;
    return {
      paragraphs: ["Send a question whenever you're ready. I can answer from the loaded knowledge documents."],
      suggestions: state.activeSuggestions
    };
  }

  if (includesAny(normalizedMessage, ["clear memory", "forget conversation", "reset memory"])) {
    clearConversationMemory();
    return {
      paragraphs: [
        "Conversation memory has been cleared.",
        "New questions will start from a clean context until you build it up again."
      ],
      suggestions: state.activeSuggestions,
      skipUserMemory: true
    };
  }

  if (includesAny(normalizedMessage, ["hello", "hi", "hey"])) {
    state.unresolvedCount = 0;
    return {
      paragraphs: [
        `Hi, I'm the ${chatbotConfig.companyName} knowledge-base assistant.`,
        "I can answer from the loaded documents, remember recent questions, and hand you over to a human when needed."
      ],
      suggestions: state.activeSuggestions
    };
  }

  if (includesAny(normalizedMessage, ["help", "what can you do", "capabilities"])) {
    state.unresolvedCount = 0;
    return {
      paragraphs: [
        "I search the loaded knowledge base, return the best matching snippets, and keep recent user questions in memory for follow-up context.",
        "Use the Load PDF / Text button to add more knowledge files at any time."
      ],
      suggestions: state.activeSuggestions
    };
  }

  if (includesAny(normalizedMessage, ["load documents", "load pdf", "upload pdf", "upload file", "add document", "how do i load"])) {
    state.unresolvedCount = 0;
    return {
      paragraphs: [
        "Use the Load PDF / Text button in the toolbar to import TXT, MD, JSON, or text-based PDF files.",
        "Once the files are loaded, I will index them and answer against the new content right away."
      ],
      suggestions: state.activeSuggestions
    };
  }

  if (includesAny(normalizedMessage, ["human", "person", "contact", "email", "phone"])) {
    state.unresolvedCount = 0;
    return {
      paragraphs: [
        `You can reach the team at ${chatbotConfig.supportEmail} or ${chatbotConfig.phone}.`,
        "If you want, use the short form below and I will capture the request for follow-up."
      ],
      suggestions: ["What services do you offer?", "How does onboarding work?"],
      showLeadForm: true,
      leadPrefillMessage: message
    };
  }

  if (!state.documents.length) {
    state.unresolvedCount = 0;
    return buildDontKnowReply({ noKnowledgeLoaded: true });
  }

  const contextualQuery = buildContextualQuery(message);
  const queryTokens = tokenize(contextualQuery.query);
  const matches = searchKnowledgeBase(contextualQuery.query);

  if (!matches.length || matches[0].score < 2.2) {
    state.unresolvedCount += 1;
    return {
      ...buildDontKnowReply({ offerHumanFollowUp: state.unresolvedCount >= 2 }),
      leadPrefillMessage: message
    };
  }

  state.unresolvedCount = 0;
  state.currentTopic = matches[0].chunk.documentTitle;
  persistConversation();
  updateMemoryStrip();

  const sourceCards = unique(matches.map((match) => match.chunk.documentTitle)).map((documentTitle) => {
    const match = matches.find((entry) => entry.chunk.documentTitle === documentTitle);
    return {
      title: match.chunk.documentTitle,
      excerpt: buildSnippet(match.chunk.text, queryTokens)
    };
  }).slice(0, 2);

  const answerLines = sourceCards.map((source) => source.excerpt);
  const introLine = contextualQuery.usedMemory
    ? "I used your recent question history to keep the context consistent."
    : "Here is the closest match from the knowledge base.";

  return {
    paragraphs: [introLine, ...answerLines],
    sources: sourceCards,
    suggestions: sourceSuggestions(matches[0].chunk.documentTitle)
  };
}

function flattenJsonValues(value, path = "", lines = []) {
  if (Array.isArray(value)) {
    value.forEach((item, index) => {
      flattenJsonValues(item, path ? `${path} item ${index + 1}` : `item ${index + 1}`, lines);
    });
    return lines;
  }

  if (value && typeof value === "object") {
    Object.entries(value).forEach(([key, nestedValue]) => {
      const nextPath = path ? `${path} ${key}` : key;
      flattenJsonValues(nestedValue, nextPath, lines);
    });
    return lines;
  }

  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    lines.push(path ? `${path}: ${value}` : String(value));
  }

  return lines;
}

function extractJsonText(rawText) {
  try {
    const parsed = JSON.parse(rawText);
    return flattenJsonValues(parsed).join("\n");
  } catch (error) {
    return rawText;
  }
}

function trimPdfStreamBytes(bytes) {
  let end = bytes.length;
  while (end > 0 && (bytes[end - 1] === 10 || bytes[end - 1] === 13)) {
    end -= 1;
  }

  return bytes.slice(0, end);
}

async function inflatePdfStream(bytes) {
  if (typeof DecompressionStream === "undefined") {
    throw new Error("This browser cannot decode compressed PDF streams.");
  }

  const decompressedStream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate"));
  const buffer = await new Response(decompressedStream).arrayBuffer();
  return new Uint8Array(buffer);
}

function readPdfLiteralString(text, startIndex) {
  let depth = 1;
  let cursor = startIndex + 1;
  let value = "";

  while (cursor < text.length && depth > 0) {
    const character = text[cursor];

    if (character === "\\") {
      const nextCharacter = text[cursor + 1];
      if (!nextCharacter) {
        cursor += 1;
        continue;
      }

      if (nextCharacter === "n" || nextCharacter === "r" || nextCharacter === "t") {
        value += " ";
        cursor += 2;
        continue;
      }

      if (/[0-7]/.test(nextCharacter)) {
        let octal = nextCharacter;
        let step = 2;
        while (step < 4 && /[0-7]/.test(text[cursor + step])) {
          octal += text[cursor + step];
          step += 1;
        }
        value += String.fromCharCode(parseInt(octal, 8));
        cursor += step;
        continue;
      }

      value += nextCharacter;
      cursor += 2;
      continue;
    }

    if (character === "(") {
      depth += 1;
      value += character;
      cursor += 1;
      continue;
    }

    if (character === ")") {
      depth -= 1;
      cursor += 1;
      if (depth === 0) {
        break;
      }
      value += ")";
      continue;
    }

    value += character;
    cursor += 1;
  }

  return {
    value: value.replace(/\s+/g, " ").trim(),
    nextIndex: cursor
  };
}

function extractPdfTextOperators(streamText) {
  const textBlocks = streamText.match(/BT[\s\S]*?ET/g) || [streamText];
  const collectedLines = [];

  textBlocks.forEach((block) => {
    const strings = [];
    let index = 0;

    while (index < block.length) {
      if (block[index] === "(") {
        const parsed = readPdfLiteralString(block, index);
        if (parsed.value) {
          strings.push(parsed.value);
        }
        index = parsed.nextIndex;
        continue;
      }

      index += 1;
    }

    if (strings.length) {
      collectedLines.push(strings.join(" "));
    }
  });

  return collectedLines
    .join("\n")
    .replace(/\s{2,}/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

async function extractPdfTextFromArrayBuffer(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const pdfText = new TextDecoder("latin1").decode(bytes);
  const streamPattern = /<<[\s\S]*?>>\s*stream\r?\n/g;
  const extractedParts = [];
  let streamMatch;

  while ((streamMatch = streamPattern.exec(pdfText)) !== null) {
    const dictionary = streamMatch[0];
    const streamStart = streamMatch.index + streamMatch[0].length;
    const streamEnd = pdfText.indexOf("endstream", streamStart);

    if (streamEnd === -1) {
      continue;
    }

    let streamBytes = trimPdfStreamBytes(bytes.slice(streamStart, streamEnd));

    if (/\/FlateDecode/.test(dictionary)) {
      try {
        streamBytes = await inflatePdfStream(streamBytes);
      } catch (error) {
        continue;
      }
    }

    if (/\/(DCTDecode|JPXDecode|CCITTFaxDecode|JBIG2Decode)/.test(dictionary)) {
      continue;
    }

    const decodedStream = new TextDecoder("latin1").decode(streamBytes);
    const extractedText = extractPdfTextOperators(decodedStream);
    if (extractedText) {
      extractedParts.push(extractedText);
    }
  }

  const combinedText = extractedParts
    .join("\n")
    .replace(/\s{2,}/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  if (combinedText) {
    return combinedText;
  }

  return extractPdfTextOperators(pdfText);
}

async function loadDocumentFromManifest(entry) {
  const type = (entry.type || "text").toLowerCase();

  if (type === "pdf") {
    const response = await fetch(entry.path, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Could not load ${entry.path}`);
    }

    const buffer = await response.arrayBuffer();
    const extractedText = await extractPdfTextFromArrayBuffer(buffer);
    return {
      id: entry.id,
      title: entry.title,
      type,
      origin: "manifest",
      text: extractedText
    };
  }

  const response = await fetch(entry.path, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`Could not load ${entry.path}`);
  }

  const rawText = await response.text();
  return {
    id: entry.id,
    title: entry.title,
    type,
    origin: "manifest",
    text: type === "json" ? extractJsonText(rawText) : rawText
  };
}

async function loadDefaultKnowledgeBase() {
  updateKnowledgeStatus("Loading knowledge base...");

  try {
    const response = await fetch(chatbotConfig.knowledgeManifest, { cache: "no-store" });
    if (!response.ok) {
      throw new Error("Manifest not found");
    }

    const manifest = await response.json();
    if (Array.isArray(manifest.suggestions) && manifest.suggestions.length) {
      state.activeSuggestions = manifest.suggestions;
    }

    const loadedDocuments = [];
    for (const entry of manifest.documents || []) {
      const documentEntry = await loadDocumentFromManifest(entry);
      if (documentEntry.text && documentEntry.text.trim()) {
        loadedDocuments.push(documentEntry);
      }
    }

    mergeKnowledgeDocuments(loadedDocuments);
    return loadedDocuments;
  } catch (error) {
    updateKnowledgeStatus("Load a PDF or text file");
    renderKnowledgeDocuments();
    return [];
  }
}

async function parseKnowledgeFile(file) {
  const extension = file.name.split(".").pop().toLowerCase();
  const id = `upload-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;

  if (["txt", "md"].includes(extension)) {
    return {
      id,
      title: humanizeFileName(file.name),
      type: extension,
      origin: "upload",
      text: await file.text()
    };
  }

  if (extension === "json") {
    return {
      id,
      title: humanizeFileName(file.name),
      type: extension,
      origin: "upload",
      text: extractJsonText(await file.text())
    };
  }

  if (extension === "pdf") {
    const extractedText = await extractPdfTextFromArrayBuffer(await file.arrayBuffer());
    if (!extractedText.trim()) {
      throw new Error("No readable text could be extracted from that PDF.");
    }

    return {
      id,
      title: humanizeFileName(file.name),
      type: extension,
      origin: "upload",
      text: extractedText
    };
  }

  throw new Error("Unsupported file type.");
}

async function handleKnowledgeUpload(files) {
  if (!files.length) {
    return;
  }

  updateKnowledgeStatus("Importing knowledge...");
  const loadedDocuments = [];
  const errors = [];

  for (const file of files) {
    try {
      const documentEntry = await parseKnowledgeFile(file);
      if (documentEntry.text.trim()) {
        loadedDocuments.push(documentEntry);
      }
    } catch (error) {
      errors.push(`${file.name}: ${error.message}`);
    }
  }

  if (loadedDocuments.length) {
    mergeKnowledgeDocuments(loadedDocuments);
    addMessage("bot", [
      `Loaded ${loadedDocuments.length} new knowledge file${loadedDocuments.length === 1 ? "" : "s"}.`,
      "You can ask questions against those documents right away."
    ]);
  }

  if (errors.length) {
    addMessage("bot", [
      "Some files could not be imported.",
      errors.join(" ")
    ]);
  }

  if (!loadedDocuments.length && !errors.length) {
    updateKnowledgeStatus("No new documents added");
  }
}

async function handleUserMessage(rawMessage) {
  const message = rawMessage.trim();
  if (!message || state.isReplyPending) {
    return;
  }

  hideLeadForm();
  addMessage("user", message);

  const reply = buildKnowledgeReply(message);

  if (!reply.skipUserMemory) {
    rememberTurn("user", message, [], {
      topicId: reply.contextTopicId || state.activeTopic || state.lastResolvedTopic
    });
  }

  setChatBusy(true);
  showTypingIndicator();

  try {
    await wait(getTypingDelay(reply));
    removeTypingIndicator();
    addMessage("bot", reply);

    if (reply.showLeadForm) {
      showLeadForm(reply.leadPrefillMessage || message);
    }

    rememberTurn("assistant", (reply.paragraphs || []).join(" "), reply.sources || []);
    setSuggestions(Array.isArray(reply.suggestions) ? reply.suggestions : state.activeSuggestions);
  } finally {
    removeTypingIndicator();
    setChatBusy(false);
    chatInput.focus();
  }
}

function seedConversation() {
  addMessage("bot", [
    `Welcome to ${chatbotConfig.companyName}.`,
    "I can answer from the loaded knowledge base, remember recent questions for follow-ups, and point you to a human when needed."
  ]);
  setSuggestions(state.activeSuggestions);
}

async function initializeChatbot() {
  updateMemoryStrip();
  renderKnowledgeDocuments();
  seedConversation();

  const loadedDocuments = await loadDefaultKnowledgeBase();
  setSuggestions(state.activeSuggestions);

  if (loadedDocuments.length) {
    addMessage("bot", [
      `I loaded ${loadedDocuments.length} starter knowledge document${loadedDocuments.length === 1 ? "" : "s"}.`,
      "Ask about services, onboarding, pricing, support, chatbot memory, or upload your own PDF and text files."
    ]);
  } else {
    addMessage("bot", [
      "No default knowledge documents were found in this environment.",
      "Use the Load PDF / Text button above to import your company knowledge base."
    ]);
  }
}

function persistConversation() {
  saveStoredJson(chatbotConfig.storageKeys.conversation, state.conversation);
  saveStoredJson(chatbotConfig.storageKeys.currentTopic, state.currentTopic);
  saveStoredJson(chatbotConfig.storageKeys.lastResolvedTopic, state.lastResolvedTopic);
  saveStoredJson(chatbotConfig.storageKeys.activeTopic, state.activeTopic);
}

function clearConversationMemory() {
  state.conversation = [];
  state.currentTopic = "";
  state.lastResolvedTopic = "";
  state.unresolvedCount = 0;
  persistConversation();
  updateMemoryStrip();
}

function getRecentConversationSignals() {
  const recentTurns = state.conversation.slice(-8);
  const recentUserTurns = recentTurns
    .filter((turn) => turn.role === "user")
    .slice(-4)
    .map((turn) => turn.text);
  const recentAssistantTurns = recentTurns
    .filter((turn) => turn.role === "assistant")
    .slice(-2)
    .map((turn) => turn.text);
  const recentSourceTitles = unique(
    recentTurns.flatMap((turn) => (Array.isArray(turn.sources) ? turn.sources : []))
  ).slice(-4);
  const recentTopicIds = unique(
    recentTurns
      .map((turn) => getOptionalNormalizedTopicId(turn.topicId))
      .filter(Boolean)
  ).slice(-3);
  const lastTopicTurn = [...recentTurns]
    .reverse()
    .find((turn) => getOptionalNormalizedTopicId(turn.topicId));

  return {
    recentUserTurns,
    recentAssistantTurns,
    recentSourceTitles,
    recentTopicIds,
    lastTopicId: lastTopicTurn ? getOptionalNormalizedTopicId(lastTopicTurn.topicId) : ""
  };
}

function buildContextualQuery(message) {
  const normalizedMessage = normalize(message);
  const explicitTopicId = detectTopicFromText(message);
  const contextSignals = getRecentConversationSignals();
  const useMemory = (
    isFollowUpQuestion(normalizedMessage) ||
    (!explicitTopicId && contextSignals.lastTopicId && normalizedMessage.split(" ").filter(Boolean).length <= 14)
  ) && state.conversation.length > 0;
  const parts = [message];

  if (useMemory) {
    contextSignals.recentUserTurns.slice(-3).forEach((question) => {
      parts.push(question);
    });

    contextSignals.recentAssistantTurns.slice(-1).forEach((answer) => {
      parts.push(answer);
    });

    contextSignals.recentSourceTitles.forEach((sourceTitle) => {
      parts.push(sourceTitle);
    });

    contextSignals.recentTopicIds.forEach((topicId) => {
      parts.push(getTopicLabel(topicId));
    });

    if (state.currentTopic) {
      parts.push(state.currentTopic);
    }

    if (state.lastResolvedTopic) {
      parts.push(getTopicLabel(state.lastResolvedTopic));
    }
  }

  return {
    query: unique(parts).join(" "),
    usedMemory: useMemory,
    explicitTopicId,
    recentTopicId: contextSignals.lastTopicId || state.lastResolvedTopic || ""
  };
}

function detectTopicFromText(text) {
  const normalizedText = normalize(text);
  const tokenSet = new Set(tokenize(text));
  let bestTopicId = "";
  let bestScore = 0;

  if (!normalizedText) {
    return "";
  }

  TOPIC_DEFINITIONS
    .filter((topicDefinition) => topicDefinition.id !== "all")
    .forEach((topicDefinition) => {
      let score = 0;

      topicDefinition.keywords.forEach((keyword) => {
        if (includesAny(normalizedText, [keyword])) {
          score += keyword.includes(" ") ? 1.5 : 1;
        }
      });

      if (tokenSet.has(topicDefinition.id)) {
        score += 2;
      }

      state.documents
        .filter((documentEntry) => documentEntry.topic === topicDefinition.id)
        .forEach((documentEntry) => {
          const normalizedTitle = normalize(documentEntry.title);
          if (normalizedTitle && normalizedText.includes(normalizedTitle)) {
            score += 2;
          }
        });

      if (score > bestScore) {
        bestScore = score;
        bestTopicId = topicDefinition.id;
      }
    });

  return bestScore > 0 ? bestTopicId : "";
}

function resolveTopicSelection(message, { usedMemory = false, explicitTopicId = "", recentTopicId = "" } = {}) {
  if (state.activeTopic !== "all") {
    return {
      topicId: state.activeTopic,
      source: "manual"
    };
  }

  const detectedTopic = explicitTopicId || detectTopicFromText(message);
  if (detectedTopic) {
    return {
      topicId: detectedTopic,
      source: "auto"
    };
  }

  if (usedMemory && recentTopicId) {
    return {
      topicId: recentTopicId,
      source: "memory"
    };
  }

  if (usedMemory && state.lastResolvedTopic) {
    return {
      topicId: state.lastResolvedTopic,
      source: "memory"
    };
  }

  return {
    topicId: "all",
    source: "all"
  };
}

function rebuildKnowledgeIndex() {
  const chunks = [];

  state.documents.forEach((documentEntry) => {
    splitIntoChunks(documentEntry.text).forEach((chunkText, index) => {
      const tokens = tokenize(`${documentEntry.title} ${getTopicLabel(documentEntry.topic)} ${chunkText}`);
      if (!tokens.length) {
        return;
      }

      chunks.push({
        id: `${documentEntry.id}-${index + 1}`,
        documentId: documentEntry.id,
        documentTitle: documentEntry.title,
        documentType: documentEntry.type,
        topic: documentEntry.topic,
        origin: documentEntry.origin,
        text: chunkText,
        normalized: normalize(`${documentEntry.title} ${chunkText}`),
        tokenCounts: countTokens(tokens),
        uniqueTokens: unique(tokens)
      });
    });
  });

  const documentFrequency = {};
  chunks.forEach((chunk) => {
    chunk.uniqueTokens.forEach((token) => {
      documentFrequency[token] = (documentFrequency[token] || 0) + 1;
    });
  });

  const totalChunks = chunks.length || 1;
  state.inverseDocumentFrequency = Object.fromEntries(
    Object.entries(documentFrequency).map(([token, frequency]) => {
      return [token, Math.log((1 + totalChunks) / (1 + frequency)) + 1];
    })
  );
  state.chunks = chunks;

  renderKnowledgeDocuments();
  renderTopicSummary();
  renderAdminControlPanel();
  renderAdminDocumentList();
  renderTopicFilters();

  if (!state.documents.length) {
    updateKnowledgeStatus("No knowledge loaded");
    return;
  }

  const topicCount = Object.values(countDocumentsByTopic()).filter(Boolean).length;
  updateKnowledgeStatus(`${state.documents.length} documents ready across ${topicCount} topics`);
}

function mergeKnowledgeDocuments(newDocuments) {
  const seenKeys = new Set(
    state.documents.map((documentEntry) => `${documentEntry.title}::${documentEntry.text}::${documentEntry.topic}`)
  );
  const addedDocuments = [];

  newDocuments
    .map((documentEntry) => prepareDocumentEntry(documentEntry))
    .forEach((documentEntry) => {
      if (!documentEntry.text) {
        return;
      }

      const key = `${documentEntry.title}::${documentEntry.text}::${documentEntry.topic}`;
      if (!seenKeys.has(key)) {
        state.documents.push(documentEntry);
        addedDocuments.push(documentEntry);
        seenKeys.add(key);
      }
    });

  rebuildKnowledgeIndex();
  return addedDocuments;
}

function scoreChunk(chunk, queryTokens, normalizedQuery, topicId = "all") {
  let score = 0;

  queryTokens.forEach((token) => {
    const tokenFrequency = chunk.tokenCounts[token] || 0;
    if (tokenFrequency) {
      score += tokenFrequency * (state.inverseDocumentFrequency[token] || 1);
    }
  });

  if (normalizedQuery && chunk.normalized.includes(normalizedQuery)) {
    score += 6;
  }

  if (topicId !== "all" && chunk.topic === topicId) {
    score += 1.8;
  }

  return score;
}

function searchKnowledgeBase(query, { topicId = "all" } = {}) {
  const normalizedQuery = normalize(query);
  const queryTokens = tokenize(query);
  const normalizedTopicId = sanitizeTopicId(topicId);

  if (!queryTokens.length || !state.chunks.length) {
    return [];
  }

  const candidateChunks = normalizedTopicId === "all"
    ? state.chunks
    : state.chunks.filter((chunk) => chunk.topic === normalizedTopicId);

  return candidateChunks
    .map((chunk) => {
      return {
        chunk,
        score: scoreChunk(chunk, queryTokens, normalizedQuery, normalizedTopicId)
      };
    })
    .filter((match) => match.score > 0)
    .sort((left, right) => right.score - left.score)
    .slice(0, 4);
}

function sourceSuggestions(topicId, documentTitle = "") {
  const normalizedTopicId = sanitizeTopicId(topicId);
  if (normalizedTopicId !== "all") {
    return getTopicSuggestions(normalizedTopicId);
  }

  const normalizedTitle = normalize(documentTitle);
  const detectedTopic = detectTopicFromText(normalizedTitle);
  return detectedTopic ? getTopicSuggestions(detectedTopic) : getDefaultQuickReplies();
}

function buildDontKnowReply({ noKnowledgeLoaded = false, offerHumanFollowUp = false, topicId = "all", manualFocus = false } = {}) {
  const normalizedTopicId = sanitizeTopicId(topicId);
  const topicLabel = normalizedTopicId === "all" ? "the current knowledge base" : `the ${getTopicLabel(normalizedTopicId)} knowledge base`;

  if (noKnowledgeLoaded) {
    return {
      paragraphs: [
        "I don't know yet because there isn't a knowledge base loaded.",
        "Use the admin section to add topic-based documents or load a temporary session file in the chat toolbar."
      ],
      suggestions: [
        "What can you do?",
        "How do I load documents?",
        "Can I talk to a human?"
      ]
    };
  }

  if (offerHumanFollowUp) {
    return {
      paragraphs: [
        `I don't know based on ${topicLabel} right now.`,
        manualFocus
          ? "You can switch the chat focus to another topic, upload a better source document, or use the follow-up form below for a human response."
          : "I've opened the follow-up form below so a teammate can help, or you can upload another document that covers this topic."
      ],
      suggestions: manualFocus
        ? getTopicSuggestions(normalizedTopicId)
        : [
            "Can I talk to a human?",
            "How do I load documents?",
            "What services do you offer?"
          ],
      showLeadForm: true
    };
  }

  return {
    paragraphs: [
      `I don't know based on ${topicLabel} right now.`,
      manualFocus
        ? "Try switching the chat focus to another topic or upload another document that covers this area."
        : "Try a more specific phrase from the source material, or upload another PDF or text file that covers this topic."
    ],
    suggestions: manualFocus ? getTopicSuggestions(normalizedTopicId) : getDefaultQuickReplies()
  };
}

function buildKnowledgeReply(message) {
  const normalizedMessage = normalize(message);
  const detectedTopicId = detectTopicFromText(message);
  const fallbackTopicId = getOptionalNormalizedTopicId(state.activeTopic) || detectedTopicId || state.lastResolvedTopic || "";

  if (!normalizedMessage) {
    state.unresolvedCount = 0;
    return {
      paragraphs: ["Send a question whenever you're ready. I can answer from the loaded knowledge documents."],
      suggestions: getDefaultQuickReplies(),
      contextTopicId: fallbackTopicId
    };
  }

  if (includesAny(normalizedMessage, ["clear memory", "forget conversation", "reset memory"])) {
    clearConversationMemory();
    return {
      paragraphs: [
        "Conversation memory has been cleared.",
        "New questions will start from a clean context until you build it up again."
      ],
      suggestions: getDefaultQuickReplies(),
      skipUserMemory: true,
      contextTopicId: ""
    };
  }

  if (includesAny(normalizedMessage, ["hello", "hi", "hey"])) {
    state.unresolvedCount = 0;
    return {
      paragraphs: [
        `Hi, I'm the ${chatbotConfig.companyName} knowledge-base assistant.`,
        "I can answer across product, support, HR, and general company topics, remember recent questions, and hand you over to a human when needed."
      ],
      suggestions: getDefaultQuickReplies(),
      contextTopicId: fallbackTopicId
    };
  }

  if (includesAny(normalizedMessage, ["help", "what can you do", "capabilities"])) {
    state.unresolvedCount = 0;
    return {
      paragraphs: [
        "I search the loaded knowledge base, route questions across topics like product, support, and HR, return the best matching snippets, and keep recent questions in memory for follow-up context.",
        "Use the admin section for persistent knowledge updates, or use Load PDF / Text in the chat for a temporary session upload."
      ],
      suggestions: getDefaultQuickReplies(),
      contextTopicId: fallbackTopicId
    };
  }

  if (includesAny(normalizedMessage, ["load documents", "load pdf", "upload pdf", "upload file", "add document", "how do i load"])) {
    state.unresolvedCount = 0;
    return {
      paragraphs: [
        "Use the admin section to upload persistent topic-based documents for product, support, HR, or general content.",
        "If you only want a temporary test file for the current browser session, use the Load PDF / Text button in the chat toolbar."
      ],
      suggestions: getDefaultQuickReplies(),
      contextTopicId: fallbackTopicId
    };
  }

  if (includesAny(normalizedMessage, ["human", "person", "contact", "email", "phone"])) {
    state.unresolvedCount = 0;
    return {
      paragraphs: [
        `You can reach the team at ${chatbotConfig.supportEmail} or ${chatbotConfig.phone}.`,
        "If you want, use the short form below and I will capture the request for follow-up."
      ],
      suggestions: getDefaultQuickReplies(),
      showLeadForm: true,
      leadPrefillMessage: message,
      contextTopicId: fallbackTopicId
    };
  }

  if (!state.documents.length) {
    state.unresolvedCount = 0;
    return {
      ...buildDontKnowReply({ noKnowledgeLoaded: true, topicId: state.activeTopic, manualFocus: state.activeTopic !== "all" }),
      contextTopicId: fallbackTopicId
    };
  }

  const contextualQuery = buildContextualQuery(message);
  const queryTokens = tokenize(contextualQuery.query);
  const topicSelection = resolveTopicSelection(message, contextualQuery);
  let matches = searchKnowledgeBase(contextualQuery.query, { topicId: topicSelection.topicId });
  let usedBroaderSearch = false;

  if (
    topicSelection.topicId !== "all" &&
    (!matches.length || matches[0].score < 2.2) &&
    topicSelection.source !== "manual"
  ) {
    const broaderMatches = searchKnowledgeBase(contextualQuery.query, { topicId: "all" });
    if (broaderMatches.length && broaderMatches[0].score >= 2.2) {
      matches = broaderMatches;
      usedBroaderSearch = true;
    }
  }

  if (!matches.length || matches[0].score < 2.2) {
    state.unresolvedCount += 1;
    return {
      ...buildDontKnowReply({
        offerHumanFollowUp: state.unresolvedCount >= 2,
        topicId: topicSelection.topicId,
        manualFocus: topicSelection.source === "manual"
      }),
      leadPrefillMessage: message,
      contextTopicId: topicSelection.topicId !== "all"
        ? topicSelection.topicId
        : contextualQuery.recentTopicId || fallbackTopicId
    };
  }

  state.unresolvedCount = 0;
  state.currentTopic = matches[0].chunk.documentTitle;
  state.lastResolvedTopic = matches[0].chunk.topic;
  persistConversation();
  updateMemoryStrip();

  const sourceCards = unique(matches.map((match) => `${match.chunk.documentTitle}::${match.chunk.topic}`))
    .map((sourceKey) => {
      const match = matches.find((entry) => `${entry.chunk.documentTitle}::${entry.chunk.topic}` === sourceKey);
      return {
        title: `${match.chunk.documentTitle} | ${getTopicLabel(match.chunk.topic)}`,
        excerpt: buildSnippet(match.chunk.text, queryTokens)
      };
    })
    .slice(0, 2);

  const answerLines = sourceCards.map((source) => source.excerpt);
  let introLine = "Here is the closest match from the knowledge base.";

  if (usedBroaderSearch) {
    introLine = `I widened the search across all topics and found the closest match in ${getTopicLabel(matches[0].chunk.topic)}.`;
  } else if (topicSelection.source === "manual" && topicSelection.topicId !== "all") {
    introLine = `Here is the closest match from the ${getTopicLabel(topicSelection.topicId)} knowledge base.`;
  } else if (topicSelection.source === "auto" && matches[0].chunk.topic !== "general") {
    introLine = `I routed this question to ${getTopicLabel(matches[0].chunk.topic)} and found the closest match there.`;
  } else if (topicSelection.source === "memory" && state.lastResolvedTopic) {
    introLine = `I kept the conversation in ${getTopicLabel(state.lastResolvedTopic)} based on your recent context.`;
  }

  return {
    paragraphs: [introLine, ...answerLines],
    sources: sourceCards,
    suggestions: sourceSuggestions(matches[0].chunk.topic, matches[0].chunk.documentTitle),
    contextTopicId: matches[0].chunk.topic
  };
}

async function loadDocumentFromManifest(entry) {
  const type = (entry.type || "text").toLowerCase();

  if (type === "pdf") {
    const response = await fetch(entry.path, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Could not load ${entry.path}`);
    }

    const buffer = await response.arrayBuffer();
    const extractedText = await extractPdfTextFromArrayBuffer(buffer);
    return prepareDocumentEntry({
      id: entry.id,
      title: entry.title,
      type,
      topic: entry.topic,
      origin: "manifest",
      text: extractedText
    });
  }

  const response = await fetch(entry.path, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`Could not load ${entry.path}`);
  }

  const rawText = await response.text();
  return prepareDocumentEntry({
    id: entry.id,
    title: entry.title,
    type,
    topic: entry.topic,
    origin: "manifest",
    text: type === "json" ? extractJsonText(rawText) : rawText
  });
}

function loadAdminKnowledgeBase() {
  const adminDocuments = getStoredAdminDocuments();
  if (adminDocuments.length) {
    mergeKnowledgeDocuments(adminDocuments);
  }

  return adminDocuments;
}

async function loadDefaultKnowledgeBase() {
  updateKnowledgeStatus("Loading knowledge base...");

  try {
    const response = await fetch(chatbotConfig.knowledgeManifest, { cache: "no-store" });
    if (!response.ok) {
      throw new Error("Manifest not found");
    }

    const manifest = await response.json();
    if (Array.isArray(manifest.suggestions) && manifest.suggestions.length) {
      state.activeSuggestions = manifest.suggestions;
    }

    const loadedDocuments = [];
    for (const entry of manifest.documents || []) {
      const documentEntry = await loadDocumentFromManifest(entry);
      if (documentEntry.text) {
        loadedDocuments.push(documentEntry);
      }
    }

    return mergeKnowledgeDocuments(loadedDocuments);
  } catch (error) {
    updateKnowledgeStatus("Load a PDF or text file");
    renderKnowledgeDocuments();
    return [];
  }
}

async function parseKnowledgeFile(file, options = {}) {
  const extension = file.name.split(".").pop().toLowerCase();
  const topicId = sanitizeTopicId(options.topic, { allowAll: false });
  const origin = options.origin || "session";
  const idPrefix = origin === "admin" ? "admin" : "upload";

  if (["txt", "md"].includes(extension)) {
    return prepareDocumentEntry({
      id: createDocumentId(idPrefix),
      title: humanizeFileName(file.name),
      type: extension,
      origin,
      topic: topicId,
      text: await file.text()
    });
  }

  if (extension === "json") {
    return prepareDocumentEntry({
      id: createDocumentId(idPrefix),
      title: humanizeFileName(file.name),
      type: extension,
      origin,
      topic: topicId,
      text: extractJsonText(await file.text())
    });
  }

  if (extension === "pdf") {
    const extractedText = await extractPdfTextFromArrayBuffer(await file.arrayBuffer());
    if (!extractedText.trim()) {
      throw new Error("No readable text could be extracted from that PDF.");
    }

    return prepareDocumentEntry({
      id: createDocumentId(idPrefix),
      title: humanizeFileName(file.name),
      type: extension,
      origin,
      topic: topicId,
      text: extractedText
    });
  }

  throw new Error("Unsupported file type.");
}

function summarizeDocumentTopics(documents) {
  return unique(documents.map((documentEntry) => getTopicLabel(documentEntry.topic))).join(", ");
}

async function handleKnowledgeUpload(files) {
  if (!files.length) {
    return;
  }

  updateKnowledgeStatus("Importing session knowledge...");
  const loadedDocuments = [];
  const errors = [];

  for (const file of files) {
    try {
      const inferredTopic = state.activeTopic !== "all"
        ? state.activeTopic
        : detectTopicFromText(file.name) || DEFAULT_TOPIC_ID;
      const documentEntry = await parseKnowledgeFile(file, {
        origin: "session",
        topic: inferredTopic
      });

      if (documentEntry.text) {
        loadedDocuments.push(documentEntry);
      }
    } catch (error) {
      errors.push(`${file.name}: ${error.message}`);
    }
  }

  if (loadedDocuments.length) {
    const addedDocuments = mergeKnowledgeDocuments(loadedDocuments);

    if (addedDocuments.length) {
      addMessage("bot", [
        `Loaded ${addedDocuments.length} new session knowledge file${addedDocuments.length === 1 ? "" : "s"} across ${summarizeDocumentTopics(addedDocuments)}.`,
        "You can ask questions against those documents right away."
      ]);
    } else {
      addMessage("bot", [
        "Those session documents were already in the knowledge base.",
        "No duplicate documents were added."
      ]);
    }
  }

  if (errors.length) {
    addMessage("bot", [
      "Some files could not be imported.",
      errors.join(" ")
    ]);
  }

  if (!loadedDocuments.length && !errors.length) {
    updateKnowledgeStatus("No new documents added");
  }
}

async function handleAdminKnowledgeUpload(files) {
  if (!files.length) {
    return;
  }

  const selectedTopic = sanitizeTopicId(adminTopicSelectEl.value, { allowAll: false });
  setAdminStatus(`Uploading ${files.length} file${files.length === 1 ? "" : "s"} to ${getTopicLabel(selectedTopic)}...`);

  const loadedDocuments = [];
  const errors = [];

  for (const file of files) {
    try {
      const documentEntry = await parseKnowledgeFile(file, {
        origin: "admin",
        topic: selectedTopic
      });

      if (documentEntry.text) {
        loadedDocuments.push(documentEntry);
      }
    } catch (error) {
      errors.push(`${file.name}: ${error.message}`);
    }
  }

  if (loadedDocuments.length) {
    const addedDocuments = mergeKnowledgeDocuments(loadedDocuments);
    saveAdminDocuments();

    if (addedDocuments.length) {
      setSelectedAdminDocument(addedDocuments[addedDocuments.length - 1].id);
      setAdminStatus(`Added ${addedDocuments.length} ${getTopicLabel(selectedTopic)} admin document${addedDocuments.length === 1 ? "" : "s"}.`);
      setAdminEditorStatus(`Selected ${addedDocuments[addedDocuments.length - 1].title} for review and updates.`);
      addMessage("bot", [
        `Knowledge base updated with ${addedDocuments.length} ${getTopicLabel(selectedTopic)} admin document${addedDocuments.length === 1 ? "" : "s"}.`,
        "I can use that content immediately."
      ]);
    } else {
      setAdminStatus("Those admin documents were already in the knowledge base.");
    }
  }

  if (errors.length) {
    setAdminStatus("Some admin uploads could not be imported.");
    addMessage("bot", [
      "Some admin documents could not be imported.",
      errors.join(" ")
    ]);
  }

  if (!loadedDocuments.length && !errors.length) {
    setAdminStatus("No admin documents were added.");
  }
}

function removeAdminDocument(documentId) {
  const documentToRemove = state.documents.find((documentEntry) => documentEntry.id === documentId && documentEntry.origin === "admin");
  if (!documentToRemove) {
    return;
  }

  const removedSelectedDocument = state.selectedAdminDocumentId === documentId;
  state.documents = state.documents.filter((documentEntry) => documentEntry.id !== documentId);
  if (removedSelectedDocument) {
    state.selectedAdminDocumentId = getAdminDocuments()[0]?.id || "";
  }
  rebuildKnowledgeIndex();
  saveAdminDocuments();
  setAdminStatus(`Removed ${documentToRemove.title} from the admin knowledge library.`);

  if (state.selectedAdminDocumentId) {
    const selectedDocument = getSelectedAdminDocument();
    if (selectedDocument) {
      setAdminEditorStatus(`${selectedDocument.title} is ready to edit. ${formatAdminTimestamp(selectedDocument.updatedAt || selectedDocument.createdAt)}.`);
    }
    return;
  }

  setAdminEditorStatus(
    getAdminDocuments().length
      ? "Select an uploaded document to update it."
      : "Upload a document to start managing the admin knowledge base."
  );
}

function clearAdminDocuments() {
  const adminDocuments = state.documents.filter((documentEntry) => documentEntry.origin === "admin");
  if (!adminDocuments.length) {
    setAdminStatus("There are no admin uploads to remove.");
    return;
  }

  state.documents = state.documents.filter((documentEntry) => documentEntry.origin !== "admin");
  state.selectedAdminDocumentId = "";
  rebuildKnowledgeIndex();
  saveAdminDocuments();
  setAdminStatus(`Removed ${adminDocuments.length} admin upload${adminDocuments.length === 1 ? "" : "s"}.`);
  setAdminEditorStatus("Upload a document to start managing the admin knowledge base.");
}

async function handleUserMessage(rawMessage) {
  const message = rawMessage.trim();
  if (!message || state.isReplyPending) {
    return;
  }

  hideLeadForm();
  addMessage("user", message);

  const reply = buildKnowledgeReply(message);

  if (!reply.skipUserMemory) {
    rememberTurn("user", message, [], {
      topicId: reply.contextTopicId || state.activeTopic || state.lastResolvedTopic
    });
  }

  setChatBusy(true);
  showTypingIndicator();

  try {
    await wait(getTypingDelay(reply));
    removeTypingIndicator();
    addMessage("bot", reply);

    if (reply.showLeadForm) {
      showLeadForm(reply.leadPrefillMessage || message);
    }

    rememberTurn("assistant", (reply.paragraphs || []).join(" "), reply.sources || [], {
      topicId: reply.contextTopicId || state.lastResolvedTopic
    });
    setSuggestions(Array.isArray(reply.suggestions) ? reply.suggestions : getDefaultQuickReplies());
  } finally {
    removeTypingIndicator();
    setChatBusy(false);
    renderTopicFilters();
    chatInput.focus();
  }
}

function seedConversation() {
  addMessage("bot", [
    `Welcome to ${chatbotConfig.companyName}.`,
    "I can answer across general, product, support, and HR knowledge, remember recent questions for follow-ups, and point you to a human when needed."
  ]);
  setSuggestions(getDefaultQuickReplies());
}

async function initializeChatbot() {
  updateMemoryStrip();
  renderTopicFilters();
  renderKnowledgeDocuments();
  renderTopicSummary();
  renderAdminControlPanel();
  renderAdminDocumentList();
  seedConversation();

  const loadedDocuments = await loadDefaultKnowledgeBase();
  const adminDocuments = loadAdminKnowledgeBase();
  setSuggestions(getDefaultQuickReplies());

  if (adminDocuments.length) {
    setAdminStatus(`${adminDocuments.length} saved admin upload${adminDocuments.length === 1 ? "" : "s"} loaded from this browser.`);
  } else {
    setAdminStatus("Starter documents are loaded. Admin uploads will persist in this browser.");
  }

  if (loadedDocuments.length || adminDocuments.length) {
    addMessage("bot", [
      `I loaded ${loadedDocuments.length} starter knowledge document${loadedDocuments.length === 1 ? "" : "s"}${adminDocuments.length ? ` and ${adminDocuments.length} saved admin upload${adminDocuments.length === 1 ? "" : "s"}` : ""}.`,
      "Ask about product behavior, support policy, HR guidance, pricing, or upload more topic-based content from the admin section."
    ]);
  } else {
    addMessage("bot", [
      "No default knowledge documents were found in this environment.",
      "Use the admin section or the Load PDF / Text button above to import your company knowledge base."
    ]);
  }
}

chatForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const message = chatInput.value;
  if (!message.trim()) {
    return;
  }

  chatInput.value = "";
  await handleUserMessage(message);
});

leadForm.addEventListener("submit", (event) => {
  event.preventDefault();

  const payload = {
    name: document.getElementById("lead-name").value.trim(),
    email: document.getElementById("lead-email").value.trim(),
    message: leadMessage.value.trim(),
    createdAt: new Date().toISOString()
  };

  const existing = loadStoredJson(chatbotConfig.storageKeys.leads, []);
  existing.push(payload);
  saveStoredJson(chatbotConfig.storageKeys.leads, existing);

  leadStatus.textContent = "Thanks. Your request has been captured locally for follow-up.";
  addMessage("bot", [
    "Thanks for sharing that.",
    `A teammate can follow up using ${payload.email}.`
  ]);
  state.unresolvedCount = 0;
  leadForm.reset();
});

chatToggle.addEventListener("click", () => {
  const isHidden = chatbotPanel.classList.contains("is-hidden");
  if (isHidden) {
    openChat();
    return;
  }

  closeChat();
});

chatClose.addEventListener("click", closeChat);
openChatbotButton.addEventListener("click", openChat);

uploadKnowledgeTriggerEl.addEventListener("click", () => {
  knowledgeUploadEl.click();
});

adminUploadTriggerEl.addEventListener("click", () => {
  adminKnowledgeUploadEl.click();
});

knowledgeUploadEl.addEventListener("change", async (event) => {
  const files = Array.from(event.target.files || []);
  await handleKnowledgeUpload(files);
  knowledgeUploadEl.value = "";
});

adminKnowledgeUploadEl.addEventListener("change", async (event) => {
  const files = Array.from(event.target.files || []);
  await handleAdminKnowledgeUpload(files);
  adminKnowledgeUploadEl.value = "";
});

adminResetDocsEl.addEventListener("click", () => {
  if (!window.confirm("Remove all admin-uploaded knowledge documents from this browser?")) {
    return;
  }

  clearAdminDocuments();
});

adminUploadedListEl.addEventListener("click", (event) => {
  const actionButton = event.target.closest("[data-admin-action]");
  if (!actionButton) {
    return;
  }

  const { adminAction, documentId } = actionButton.dataset;
  if (!documentId) {
    return;
  }

  if (adminAction === "edit") {
    setSelectedAdminDocument(documentId);
    return;
  }

  if (adminAction === "delete") {
    const selectedDocument = state.documents.find((documentEntry) => {
      return documentEntry.origin === "admin" && documentEntry.id === documentId;
    });

    if (!selectedDocument) {
      return;
    }

    if (!window.confirm(`Delete ${selectedDocument.title} from the admin knowledge base?`)) {
      return;
    }

    removeAdminDocument(documentId);
  }
});

adminEditorFormEl.addEventListener("submit", (event) => {
  event.preventDefault();

  const documentId = adminEditorIdEl.value;
  if (!documentId) {
    setAdminEditorStatus("Select an uploaded document before saving changes.");
    return;
  }

  const title = adminEditorTitleEl.value.trim();
  const text = adminEditorContentEl.value.trim();
  if (!title) {
    setAdminEditorStatus("Add a document title before saving.");
    adminEditorTitleEl.focus();
    return;
  }

  if (!text) {
    setAdminEditorStatus("Document content cannot be empty.");
    adminEditorContentEl.focus();
    return;
  }

  setAdminEditorStatus("Saving document changes...");
  updateAdminDocument(documentId, {
    title,
    topic: adminEditorTopicEl.value,
    text
  });
});

adminCancelEditEl.addEventListener("click", () => {
  state.selectedAdminDocumentId = "";
  renderAdminControlPanel();
  setAdminEditorStatus(
    getAdminDocuments().length
      ? "Selection cleared. Choose another uploaded document to edit."
      : "Upload a document to start managing the admin knowledge base."
  );
});

clearMemoryButtonEl.addEventListener("click", () => {
  clearConversationMemory();
  hideLeadForm();
  addMessage("bot", [
    "Memory cleared.",
    "The next question will start from a fresh context."
  ]);
  setSuggestions(getDefaultQuickReplies());
});

initializeChatbot();
