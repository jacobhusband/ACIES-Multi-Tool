import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { Node, mergeAttributes } from "@tiptap/core";
import Color from "@tiptap/extension-color";
import { Details, DetailsContent, DetailsSummary } from "@tiptap/extension-details";
import Highlight from "@tiptap/extension-highlight";
import Image from "@tiptap/extension-image";
import Link from "@tiptap/extension-link";
import { TableKit } from "@tiptap/extension-table";
import TaskItem from "@tiptap/extension-task-item";
import TaskList from "@tiptap/extension-task-list";
import { TextStyle } from "@tiptap/extension-text-style";
import Underline from "@tiptap/extension-underline";
import { EditorContent, useEditor } from "@tiptap/react";
import StarterKit from "@tiptap/starter-kit";
import "./styles.css";

let root = null;
let mountedContainer = null;
let currentContext = null;
let currentOptions = {};
let flushCurrentEditor = null;
const listeners = new Set();

function emitContext() {
  listeners.forEach((listener) => listener(currentContext));
}

function subscribe(listener) {
  listeners.add(listener);
  listener(currentContext);
  return () => listeners.delete(listener);
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error || new Error("Could not read file."));
    reader.readAsDataURL(file);
  });
}

function stripTransientPageHtml(html) {
  const doc = new DOMParser().parseFromString(`<div>${html || ""}</div>`, "text/html");
  const rootEl = doc.body.firstElementChild;
  rootEl.querySelectorAll("mark.page-find-match").forEach((mark) => {
    mark.replaceWith(doc.createTextNode(mark.textContent || ""));
  });
  rootEl.querySelectorAll("img").forEach((img) => {
    img.classList.add("page-inline-image");
    img.classList.remove("is-selected");
    if (img.dataset.asset || img.getAttribute("data-asset")) {
      img.removeAttribute("src");
    }
  });
  rootEl.querySelectorAll("a.page-wiki-link").forEach((link) => {
    link.classList.remove("is-broken");
    link.removeAttribute("title");
    link.setAttribute("contenteditable", "false");
  });
  return rootEl.innerHTML;
}

function getActiveBlockRect(editor, selector) {
  const { $from } = editor.state.selection;
  const domAt = editor.view.domAtPos($from.pos);
  const element = domAt.node.nodeType === 1 ? domAt.node : domAt.node.parentElement;
  return element?.closest?.(selector)?.getBoundingClientRect() || null;
}

function detectSlashQuery(editor) {
  if (!editor) return null;
  const { state } = editor;
  const { selection } = state;
  if (!selection.empty) return null;
  const $from = selection.$from;
  const before = $from.parent.textBetween(0, $from.parentOffset, "\n", "\0");
  const match = before.match(/(?:^|\s)\/([a-z0-9 ]*)$/i);
  if (!match) return null;
  const query = match[1] || "";
  const from = $from.pos - query.length - 1;
  return {
    query,
    range: { from, to: $from.pos },
    pos: editor.view.coordsAtPos($from.pos),
  };
}

function detectPageLinkQuery(editor) {
  if (!editor) return null;
  const { state } = editor;
  const { selection } = state;
  if (!selection.empty) return null;
  const $from = selection.$from;
  const before = $from.parent.textBetween(0, $from.parentOffset, "\n", "\0");
  const match = before.match(/\[\[([^\]\[]*)$/);
  if (!match) return null;
  const query = match[1] || "";
  const from = $from.pos - query.length - 2;
  return {
    query,
    range: { from, to: $from.pos },
    pos: editor.view.coordsAtPos($from.pos),
  };
}

const PageImage = Image.extend({
  addAttributes() {
    return {
      ...this.parent?.(),
      assetPath: {
        default: null,
        parseHTML: (element) => element.getAttribute("data-asset"),
        renderHTML: (attributes) =>
          attributes.assetPath ? { "data-asset": attributes.assetPath } : {},
      },
      widthPercent: {
        default: null,
        parseHTML: (element) => element.getAttribute("data-width-percent"),
        renderHTML: (attributes) =>
          attributes.widthPercent ? { "data-width-percent": attributes.widthPercent } : {},
      },
      class: {
        default: "page-inline-image",
        parseHTML: (element) => element.getAttribute("class") || "page-inline-image",
        renderHTML: (attributes) => ({ class: attributes.class || "page-inline-image" }),
      },
      style: {
        default: null,
        parseHTML: (element) => element.getAttribute("style"),
        renderHTML: (attributes) => (attributes.style ? { style: attributes.style } : {}),
      },
    };
  },
});

const PageLink = Node.create({
  name: "pageLink",
  group: "inline",
  inline: true,
  atom: true,
  selectable: false,

  addAttributes() {
    return {
      pageId: {
        default: "",
        parseHTML: (element) => element.getAttribute("data-page-id") || "",
      },
      title: {
        default: "Untitled",
        parseHTML: (element) => element.textContent || "Untitled",
      },
    };
  },

  parseHTML() {
    return [{ tag: "a.page-wiki-link[data-page-id]" }];
  },

  renderHTML({ HTMLAttributes }) {
    const title = HTMLAttributes.title || "Untitled";
    return [
      "a",
      mergeAttributes({
        class: "page-wiki-link",
        href: "#",
        "data-page-id": HTMLAttributes.pageId,
        contenteditable: "false",
      }),
      title,
    ];
  },

  addCommands() {
    return {
      insertPageLink:
        (attrs) =>
        ({ chain }) =>
          chain().insertContent({ type: this.name, attrs }).run(),
    };
  },
});

const CALLOUT_EMOJIS = ["💡", "📌", "⚠️", "✅", "❓", "🔥"];
const CALLOUT_COLORS = ["gray", "blue", "green", "yellow", "red", "purple"];

const PageCallout = Node.create({
  name: "pageCallout",
  group: "block",
  content: "paragraph+",
  defining: true,

  addAttributes() {
    return {
      color: {
        default: "gray",
        parseHTML: (element) => element.getAttribute("data-callout-color") || "gray",
        renderHTML: (attributes) => ({ "data-callout-color": attributes.color || "gray" }),
      },
      emoji: {
        default: "💡",
        parseHTML: (element) => element.getAttribute("data-callout-emoji") || "💡",
        renderHTML: (attributes) => ({ "data-callout-emoji": attributes.emoji || "💡" }),
      },
    };
  },

  parseHTML() {
    return [{ tag: "div.page-callout" }];
  },

  renderHTML({ HTMLAttributes }) {
    return ["div", mergeAttributes({ class: "page-callout" }, HTMLAttributes), 0];
  },

  addCommands() {
    return {
      setCallout:
        (attrs = {}) =>
        ({ commands }) =>
          commands.wrapIn(this.name, attrs),
      updateCalloutAttrs:
        (attrs) =>
        ({ commands }) =>
          commands.updateAttributes(this.name, attrs),
    };
  },

  addKeyboardShortcuts() {
    return {
      "Mod-Enter": () => {
        const { state } = this.editor;
        const { $from } = state.selection;
        for (let depth = $from.depth; depth > 0; depth -= 1) {
          if ($from.node(depth).type.name === this.name) {
            const pos = $from.after(depth);
            return this.editor
              .chain()
              .insertContentAt(pos, { type: "paragraph" })
              .focus(pos + 1)
              .run();
          }
        }
        return false;
      },
    };
  },
});

const TEXT_COLORS = [
  { label: "Default", value: null },
  { label: "Gray", value: "#9b9a97" },
  { label: "Brown", value: "#a8775c" },
  { label: "Orange", value: "#d9730d" },
  { label: "Yellow", value: "#c29343" },
  { label: "Green", value: "#4d9968" },
  { label: "Blue", value: "#3f83c8" },
  { label: "Purple", value: "#9d68d3" },
  { label: "Pink", value: "#d5598f" },
  { label: "Red", value: "#e03e3e" },
];

const HIGHLIGHT_COLORS = [
  { label: "Default", value: null },
  { label: "Gray", value: "rgba(148, 163, 184, 0.35)" },
  { label: "Brown", value: "rgba(180, 130, 90, 0.35)" },
  { label: "Orange", value: "rgba(251, 146, 60, 0.35)" },
  { label: "Yellow", value: "rgba(250, 204, 21, 0.4)" },
  { label: "Green", value: "rgba(74, 222, 128, 0.35)" },
  { label: "Blue", value: "rgba(96, 165, 250, 0.35)" },
  { label: "Purple", value: "rgba(192, 132, 252, 0.35)" },
  { label: "Pink", value: "rgba(244, 114, 182, 0.35)" },
  { label: "Red", value: "rgba(248, 113, 113, 0.35)" },
];

const SLASH_COMMANDS = [
  { id: "text", label: "Text", shortcut: "P", group: "block" },
  { id: "h1", label: "Heading 1", shortcut: "H1", group: "block" },
  { id: "h2", label: "Heading 2", shortcut: "H2", group: "block" },
  { id: "h3", label: "Heading 3", shortcut: "H3", group: "block" },
  { id: "bullet", label: "Bulleted list", shortcut: "*", group: "block" },
  { id: "numbered", label: "Numbered list", shortcut: "1.", group: "block" },
  { id: "todo", label: "Todo", shortcut: "[]", group: "block" },
  { id: "quote", label: "Quote", shortcut: ">", group: "block" },
  { id: "callout", label: "Callout", shortcut: "!", group: "block" },
  { id: "toggle", label: "Toggle", shortcut: ">v", group: "block" },
  { id: "table", label: "Table", shortcut: "3x3", group: "block" },
  { id: "code", label: "Code block", shortcut: "```", group: "block" },
  { id: "divider", label: "Divider", shortcut: "---", group: "block" },
  { id: "bold", label: "Bold", shortcut: "B", group: "inline" },
  { id: "italic", label: "Italic", shortcut: "I", group: "inline" },
  { id: "underline", label: "Underline", shortcut: "U", group: "inline" },
  { id: "link", label: "Link", shortcut: "->", group: "inline" },
  { id: "color", label: "Text color", shortcut: "A", group: "inline" },
  { id: "highlight", label: "Highlight", shortcut: "==", group: "inline" },
  { id: "image", label: "Image", shortcut: "Img", group: "media" },
  { id: "page", label: "Page", shortcut: "+", group: "page", projectOnly: true },
  { id: "pageref", label: "Link to page", shortcut: "[[", group: "page" },
];

function commandMatches(command, query) {
  const q = String(query || "").trim().toLowerCase();
  if (!q) return true;
  return command.label.toLowerCase().includes(q) || command.id.includes(q);
}

function PageEditorApp() {
  const [context, setContext] = useState(currentContext);

  useEffect(() => subscribe(setContext), []);

  if (!context) return null;
  return <PageEditor context={context} options={currentOptions} />;
}

function PageEditor({ context, options }) {
  const fileInputRef = useRef(null);
  const saveTimerRef = useRef(null);
  const titleTimerRef = useRef(null);
  const suppressUpdateRef = useRef(false);
  const [title, setTitle] = useState(context.title || "");
  const [slash, setSlash] = useState({ open: false, query: "", selected: 0, range: null, pos: null });
  const [pageLink, setPageLink] = useState({ open: false, query: "", selected: 0, range: null, pos: null });
  const [colorMenu, setColorMenu] = useState({ open: false, mode: "color", pos: null });
  const [calloutMenu, setCalloutMenu] = useState({ open: false, pos: null });
  const [tableMenu, setTableMenu] = useState({ open: false, pos: null });

  const globalPages = Array.isArray(context.globalPages) ? context.globalPages : [];
  const pageLinkMatches = useMemo(() => {
    const q = String(pageLink.query || "").trim().toLowerCase();
    const matched = globalPages
      .filter((page) => !q || String(page.title || "").toLowerCase().includes(q))
      .slice(0, 8)
      .map((page) => ({ type: "page", page }));
    const exact = globalPages.some((page) => String(page.title || "").trim().toLowerCase() === q);
    if (q && !exact) matched.push({ type: "create", title: pageLink.query.trim() });
    return matched;
  }, [globalPages, pageLink.query]);

  const visibleSlashCommands = useMemo(
    () =>
      SLASH_COMMANDS.filter((command) => {
        if (command.projectOnly && context.kind === "global") return false;
        return commandMatches(command, slash.query);
      }),
    [context.kind, slash.query]
  );

  const editor = useEditor({
    extensions: [
      StarterKit.configure({
        heading: { levels: [1, 2, 3] },
        link: false,
        underline: false,
      }),
      Underline,
      TextStyle,
      Color,
      Link.configure({
        autolink: true,
        openOnClick: false,
        HTMLAttributes: { target: "_blank", rel: "noopener noreferrer" },
      }),
      TaskList,
      TaskItem.configure({ nested: true }),
      Highlight.configure({ multicolor: true }),
      TableKit.configure({ table: { resizable: false } }),
      Details.configure({ persist: true, HTMLAttributes: { class: "page-toggle" } }),
      DetailsSummary,
      DetailsContent,
      PageCallout,
      PageImage,
      PageLink,
    ],
    content: context.html || "",
    editorProps: {
      attributes: {
        id: "pageEditor",
        class: "page-editor project-pages-editor-content",
        spellcheck: "true",
        role: "textbox",
        "aria-multiline": "true",
        "aria-label": "Page content",
        "data-placeholder": "Type here - add headings, images, links, and notes...",
      },
      handleKeyDown(view, event) {
        if (colorMenu.open && event.key === "Escape") {
          event.preventDefault();
          setColorMenu((state) => ({ ...state, open: false }));
          return true;
        }
        if (pageLink.open) {
          if (event.key === "Escape") {
            event.preventDefault();
            setPageLink((state) => ({ ...state, open: false }));
            return true;
          }
          if (event.key === "ArrowDown" || event.key === "ArrowUp") {
            event.preventDefault();
            setPageLink((state) => ({
              ...state,
              selected:
                (state.selected + (event.key === "ArrowDown" ? 1 : -1) + pageLinkMatches.length) %
                Math.max(pageLinkMatches.length, 1),
            }));
            return true;
          }
          if (event.key === "Enter" || event.key === "Tab") {
            event.preventDefault();
            choosePageLinkMatch(pageLinkMatches[pageLink.selected] || pageLinkMatches[0]);
            return true;
          }
        }
        if (slash.open) {
          if (event.key === "Escape") {
            event.preventDefault();
            setSlash((state) => ({ ...state, open: false }));
            return true;
          }
          if (event.key === "ArrowDown" || event.key === "ArrowUp") {
            event.preventDefault();
            setSlash((state) => ({
              ...state,
              selected:
                (state.selected + (event.key === "ArrowDown" ? 1 : -1) + visibleSlashCommands.length) %
                Math.max(visibleSlashCommands.length, 1),
            }));
            return true;
          }
          if (event.key === "Enter" || event.key === "Tab") {
            event.preventDefault();
            executeSlashCommand(visibleSlashCommands[slash.selected] || visibleSlashCommands[0]);
            return true;
          }
        }
        if (event.key === "Tab" && !event.ctrlKey && !event.metaKey && !event.altKey) {
          if (editor?.isActive("table")) {
            event.preventDefault();
            if (event.shiftKey) editor.chain().focus().goToPreviousCell().run();
            else editor.chain().focus().goToNextCell().run();
            return true;
          }
          if (editor?.isActive("codeBlock")) {
            if (event.shiftKey) return false;
            event.preventDefault();
            editor.chain().focus().insertContent("  ").run();
            return true;
          }
          const chain = editor?.chain().focus();
          if (chain) {
            event.preventDefault();
            if (event.shiftKey) chain.liftListItem("listItem").run();
            else chain.sinkListItem("listItem").run();
            return true;
          }
        }
        return false;
      },
      handlePaste(view, event) {
        const files = Array.from(event.clipboardData?.files || []).filter((file) =>
          String(file.type || "").toLowerCase().startsWith("image/")
        );
        if (!files.length) return false;
        event.preventDefault();
        void insertImageFiles(files);
        return true;
      },
      handleDrop(view, event) {
        const files = Array.from(event.dataTransfer?.files || []).filter((file) =>
          String(file.type || "").toLowerCase().startsWith("image/")
        );
        if (!files.length) return false;
        event.preventDefault();
        void insertImageFiles(files);
        return true;
      },
    },
    onUpdate({ editor: activeEditor }) {
      if (suppressUpdateRef.current) return;
      queueHtmlSave(activeEditor.getHTML());
      refreshMenus(activeEditor);
    },
    onSelectionUpdate({ editor: activeEditor }) {
      refreshMenus(activeEditor);
    },
  });

  const hydrateImages = useCallback(async () => {
    if (!editor || !context.onGetAsset) return;
    const imgs = Array.from(editor.view.dom.querySelectorAll("img[data-asset]"));
    for (const img of imgs) {
      if (img.getAttribute("src")) continue;
      try {
        const result = await context.onGetAsset(img.getAttribute("data-asset"));
        if (result?.status === "success" && result.dataUrl) {
          img.src = result.dataUrl;
        } else {
          img.alt = "Image unavailable";
          img.classList.add("page-image-missing");
        }
      } catch (error) {
        console.warn("Page image hydrate failed:", error);
      }
    }
  }, [context, editor]);

  const flushSave = useCallback(async () => {
    if (saveTimerRef.current) {
      clearTimeout(saveTimerRef.current);
      saveTimerRef.current = null;
    }
    if (titleTimerRef.current) {
      clearTimeout(titleTimerRef.current);
      titleTimerRef.current = null;
      context.onTitleChange?.(title);
    }
    if (editor) {
      await context.onHtmlChange?.(stripTransientPageHtml(editor.getHTML()), { immediate: true });
    }
  }, [context, editor, title]);

  useEffect(() => {
    flushCurrentEditor = flushSave;
    return () => {
      if (flushCurrentEditor === flushSave) flushCurrentEditor = null;
    };
  }, [flushSave]);

  useEffect(() => {
    if (!editor) return;
    suppressUpdateRef.current = true;
    editor.commands.setContent(context.html || "", false);
    suppressUpdateRef.current = false;
    setTitle(context.title || "");
    setSlash({ open: false, query: "", selected: 0, range: null, pos: null });
    setPageLink({ open: false, query: "", selected: 0, range: null, pos: null });
    requestAnimationFrame(() => {
      void hydrateImages();
      editor.commands.focus("end");
    });
  }, [context.documentKey, editor]);

  useEffect(() => {
    void hydrateImages();
  }, [context.documentKey, hydrateImages]);

  function queueHtmlSave(html) {
    context.onHtmlChange?.(stripTransientPageHtml(html), { immediate: false });
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(() => {
      saveTimerRef.current = null;
      context.onRequestPersist?.();
    }, 700);
  }

  function refreshMenus(activeEditor) {
    const slashQuery = detectSlashQuery(activeEditor);
    if (slashQuery) {
      setSlash((state) => ({
        open: true,
        query: slashQuery.query,
        range: slashQuery.range,
        pos: slashQuery.pos,
        selected: Math.min(state.selected || 0, Math.max(visibleSlashCommands.length - 1, 0)),
      }));
    } else {
      setSlash((state) => (state.open ? { ...state, open: false } : state));
    }

    setColorMenu((state) => (state.open ? { ...state, open: false } : state));

    if (activeEditor.isActive("pageCallout")) {
      const rect = getActiveBlockRect(activeEditor, ".page-callout");
      setCalloutMenu({ open: true, pos: rect ? { left: rect.left, top: rect.top } : null });
    } else {
      setCalloutMenu((state) => (state.open ? { ...state, open: false } : state));
    }

    if (activeEditor.isActive("table")) {
      const rect = getActiveBlockRect(activeEditor, "table");
      setTableMenu({ open: true, pos: rect ? { left: rect.left, top: rect.top } : null });
    } else {
      setTableMenu((state) => (state.open ? { ...state, open: false } : state));
    }

    const pageQuery = detectPageLinkQuery(activeEditor);
    if (pageQuery) {
      setPageLink((state) => ({
        open: true,
        query: pageQuery.query,
        range: pageQuery.range,
        pos: pageQuery.pos,
        selected: Math.min(state.selected || 0, Math.max(pageLinkMatches.length - 1, 0)),
      }));
    } else {
      setPageLink((state) => (state.open ? { ...state, open: false } : state));
    }
  }

  function deleteRange(range) {
    if (!editor || !range) return editor?.chain().focus();
    return editor.chain().focus().deleteRange(range);
  }

  function executeSlashCommand(command) {
    if (!editor || !command) return;
    let chain = deleteRange(slash.range);
    if (command.id === "text") chain.setParagraph().run();
    else if (command.id === "h1") chain.toggleHeading({ level: 1 }).run();
    else if (command.id === "h2") chain.toggleHeading({ level: 2 }).run();
    else if (command.id === "h3") chain.toggleHeading({ level: 3 }).run();
    else if (command.id === "bullet") chain.toggleBulletList().run();
    else if (command.id === "numbered") chain.toggleOrderedList().run();
    else if (command.id === "todo") chain.toggleTaskList().run();
    else if (command.id === "quote") chain.toggleBlockquote().run();
    else if (command.id === "callout") chain.setCallout().run();
    else if (command.id === "toggle") chain.setDetails().run();
    else if (command.id === "table") chain.insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run();
    else if (command.id === "code") chain.toggleCodeBlock().run();
    else if (command.id === "divider") chain.setHorizontalRule().run();
    else if (command.id === "bold") chain.toggleBold().run();
    else if (command.id === "italic") chain.toggleItalic().run();
    else if (command.id === "underline") chain.toggleUnderline().run();
    else if (command.id === "color" || command.id === "highlight") {
      const pos = slash.pos;
      chain.run();
      requestAnimationFrame(() =>
        setColorMenu({ open: true, mode: command.id === "highlight" ? "highlight" : "color", pos })
      );
    }
    else if (command.id === "link") {
      chain.run();
      const url = window.prompt("Link URL", "https://");
      if (url && url.trim() && url.trim() !== "https://") {
        editor.chain().focus().extendMarkRange("link").setLink({ href: normalizeHref(url) }).run();
      }
    } else if (command.id === "image") {
      chain.run();
      fileInputRef.current?.click();
    } else if (command.id === "page") {
      chain.run();
      context.onCreateSubpage?.();
    } else if (command.id === "pageref") {
      chain.insertContent("[[").run();
    }
    setSlash((state) => ({ ...state, open: false }));
  }

  function applyColorChoice(value) {
    if (!editor) return;
    const chain = editor.chain().focus();
    if (colorMenu.mode === "highlight") {
      if (value) chain.setHighlight({ color: value }).run();
      else chain.unsetHighlight().run();
    } else {
      if (value) chain.setColor(value).run();
      else chain.unsetColor().run();
    }
    setColorMenu((state) => ({ ...state, open: false }));
  }

  function choosePageLinkMatch(match) {
    if (!editor || !match) return;
    let page = match.page;
    if (match.type === "create") {
      page = context.onCreateGlobalPage?.(match.title);
    }
    if (!page) return;
    deleteRange(pageLink.range)
      .insertPageLink({ pageId: page.id, title: page.title || "Untitled" })
      .run();
    setPageLink((state) => ({ ...state, open: false }));
  }

  async function insertImageFiles(files) {
    if (!editor || !context.onSaveAsset) return;
    for (const file of Array.from(files || [])) {
      if (!String(file.type || "").toLowerCase().startsWith("image/")) continue;
      try {
        const dataUrl = await readFileAsDataUrl(file);
        const result = await context.onSaveAsset(dataUrl, file.name || "");
        if (result?.status !== "success" || !result.assetPath) {
          context.onToast?.(result?.message || "Could not save image.");
          continue;
        }
        editor
          .chain()
          .focus()
          .setImage({
            src: dataUrl,
            assetPath: result.assetPath,
            alt: file.name || "Image",
            widthPercent: "80",
            class: "page-inline-image",
            style: "max-width: 100%;",
          })
          .run();
      } catch (error) {
        console.warn("Page image insert failed:", error);
      }
    }
  }

  function handleTitleInput(event) {
    const value = event.currentTarget.textContent || "";
    setTitle(value);
    if (titleTimerRef.current) clearTimeout(titleTimerRef.current);
    titleTimerRef.current = setTimeout(() => {
      titleTimerRef.current = null;
      context.onTitleChange?.(value);
    }, 500);
  }

  function handleEditorClick(event) {
    const pageAnchor = event.target.closest?.("a.page-wiki-link[data-page-id]");
    if (pageAnchor) {
      event.preventDefault();
      context.onOpenGlobalPage?.(pageAnchor.getAttribute("data-page-id"));
      return;
    }
    const link = event.target.closest?.("a[href]");
    if (link && (event.ctrlKey || event.metaKey)) {
      event.preventDefault();
      window.open(link.href, "_blank", "noopener,noreferrer");
    }
  }

  return (
    <div className="project-pages-editor" onClick={handleEditorClick}>
      <h1
        id="pageTitle"
        className="page-title project-pages-title"
        contentEditable
        suppressContentEditableWarning
        spellCheck
        role="textbox"
        aria-label="Page title"
        data-placeholder="Untitled"
        onInput={handleTitleInput}
        onKeyDown={(event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            editor?.commands.focus();
          }
        }}
      >
        {title}
      </h1>

      {context.kind !== "global" && (
        <div className="page-child-links" id="pageChildLinks" aria-label="Child pages">
          <div className="page-child-links-label">Subpages</div>
          {(context.childPages || []).map((child) => (
            <button
              className="page-child-link"
              type="button"
              key={child.id}
              title={child.title || "Untitled"}
              onClick={() => context.onOpenSubpage?.(child.id)}
            >
              <span className="page-child-link-icon">Pg</span>
              <span className="page-child-link-title">{child.title || "Untitled"}</span>
              <span className="page-child-link-meta">
                {child.childCount
                  ? `${child.childCount} subpage${child.childCount === 1 ? "" : "s"}`
                  : ""}
              </span>
              <span
                className="tab-delete-icon"
                role="button"
                tabIndex={0}
                title="Delete subpage"
                onClick={(event) => {
                  event.stopPropagation();
                  context.onDeleteSubpage?.(child.id);
                }}
              >
                x
              </span>
            </button>
          ))}
          <button
            className="page-child-link page-child-link-add"
            type="button"
            onClick={() => context.onCreateSubpage?.()}
          >
            + Add subpage
          </button>
        </div>
      )}

      <div className="project-pages-editor-surface">
        <EditorContent editor={editor} />
      </div>

      <input
        ref={fileInputRef}
        type="file"
        accept="image/*"
        multiple
        hidden
        onChange={(event) => {
          void insertImageFiles(event.currentTarget.files || []);
          event.currentTarget.value = "";
        }}
      />

      <CommandMenu
        id="pageSlashMenu"
        open={slash.open}
        position={slash.pos}
        header={slash.query ? `/${slash.query}` : "Type a command"}
        rows={visibleSlashCommands.map((command) => ({
          key: command.id,
          label: command.label,
          shortcut: command.shortcut,
          selected: visibleSlashCommands[slash.selected]?.id === command.id,
          onMouseDown: () => executeSlashCommand(command),
        }))}
      />
      <TableMenu
        open={tableMenu.open}
        position={tableMenu.pos}
        editor={editor}
      />
      <CalloutMenu
        open={calloutMenu.open}
        position={calloutMenu.pos}
        onEmoji={(emoji) => editor?.chain().focus().updateCalloutAttrs({ emoji }).run()}
        onColor={(color) => editor?.chain().focus().updateCalloutAttrs({ color }).run()}
      />
      <ColorMenu
        open={colorMenu.open}
        mode={colorMenu.mode}
        position={colorMenu.pos}
        colors={colorMenu.mode === "highlight" ? HIGHLIGHT_COLORS : TEXT_COLORS}
        onPick={applyColorChoice}
      />
      <CommandMenu
        id="pageLinkMenu"
        open={pageLink.open}
        position={pageLink.pos}
        header={pageLink.query ? `[[${pageLink.query}` : "Link to page"}
        rows={pageLinkMatches.map((match, index) => ({
          key: match.type === "create" ? `create:${match.title}` : match.page.id,
          label: match.type === "create" ? `Create "${match.title}"` : match.page.title || "Untitled",
          shortcut: match.type === "create" ? "New" : "Page",
          selected: index === pageLink.selected,
          onMouseDown: () => choosePageLinkMatch(match),
        }))}
      />
    </div>
  );
}

function CommandMenu({ id, open, position, header, rows }) {
  if (!open || !rows.length) return null;
  const style = position
    ? {
        left: `${Math.max(16, position.left)}px`,
        top: `${Math.max(16, position.bottom + 8)}px`,
      }
    : {};
  return (
    <div className="page-slash-menu project-pages-menu" id={id} role="listbox" style={style}>
      <div className="page-slash-menu-header">
        <span>Commands</span>
        <span>{header}</span>
      </div>
      {rows.map((row) => (
        <button
          className={`page-slash-menu-item ${row.selected ? "is-active" : ""}`}
          type="button"
          role="option"
          aria-selected={row.selected ? "true" : "false"}
          key={row.key}
          onMouseDown={(event) => {
            event.preventDefault();
            row.onMouseDown();
          }}
        >
          <span className="page-slash-menu-label">{row.label}</span>
          <span className="page-slash-menu-shortcut">{row.shortcut}</span>
        </button>
      ))}
    </div>
  );
}

const TABLE_ACTIONS = [
  { id: "rowAbove", label: "+Row ↑", title: "Add row above", run: (chain) => chain.addRowBefore() },
  { id: "rowBelow", label: "+Row ↓", title: "Add row below", run: (chain) => chain.addRowAfter() },
  { id: "colLeft", label: "+Col ←", title: "Add column left", run: (chain) => chain.addColumnBefore() },
  { id: "colRight", label: "+Col →", title: "Add column right", run: (chain) => chain.addColumnAfter() },
  { id: "delRow", label: "−Row", title: "Delete row", run: (chain) => chain.deleteRow() },
  { id: "delCol", label: "−Col", title: "Delete column", run: (chain) => chain.deleteColumn() },
  { id: "header", label: "Header", title: "Toggle header row", run: (chain) => chain.toggleHeaderRow() },
  { id: "delTable", label: "✕", title: "Delete table", run: (chain) => chain.deleteTable() },
];

function TableMenu({ open, position, editor }) {
  if (!open || !editor) return null;
  const style = position
    ? {
        left: `${Math.max(16, position.left)}px`,
        top: `${Math.max(16, position.top - 44)}px`,
      }
    : {};
  return (
    <div className="page-slash-menu project-pages-menu project-pages-table-menu" style={style}>
      {TABLE_ACTIONS.map((action) => (
        <button
          key={action.id}
          type="button"
          className="project-pages-table-btn"
          title={action.title}
          aria-label={action.title}
          onMouseDown={(event) => {
            event.preventDefault();
            action.run(editor.chain().focus()).run();
          }}
        >
          {action.label}
        </button>
      ))}
    </div>
  );
}

function CalloutMenu({ open, position, onEmoji, onColor }) {
  if (!open) return null;
  const style = position
    ? {
        left: `${Math.max(16, position.left)}px`,
        top: `${Math.max(16, position.top - 44)}px`,
      }
    : {};
  return (
    <div className="page-slash-menu project-pages-menu project-pages-callout-menu" style={style}>
      {CALLOUT_EMOJIS.map((emoji) => (
        <button
          key={emoji}
          type="button"
          className="project-pages-callout-btn"
          title="Callout icon"
          onMouseDown={(event) => {
            event.preventDefault();
            onEmoji(emoji);
          }}
        >
          {emoji}
        </button>
      ))}
      <span className="project-pages-callout-sep" />
      {CALLOUT_COLORS.map((color) => (
        <button
          key={color}
          type="button"
          className={`project-pages-callout-dot callout-dot-${color}`}
          title={`${color} callout`}
          aria-label={`${color} callout`}
          onMouseDown={(event) => {
            event.preventDefault();
            onColor(color);
          }}
        />
      ))}
    </div>
  );
}

function ColorMenu({ open, mode, position, colors, onPick }) {
  if (!open) return null;
  const style = position
    ? {
        left: `${Math.max(16, position.left)}px`,
        top: `${Math.max(16, position.bottom + 8)}px`,
      }
    : {};
  return (
    <div className="page-slash-menu project-pages-menu project-pages-color-menu" role="listbox" style={style}>
      <div className="page-slash-menu-header">
        <span>{mode === "highlight" ? "Highlight" : "Text color"}</span>
        <span>Esc to close</span>
      </div>
      <div className="project-pages-color-grid">
        {colors.map((color) => (
          <button
            key={color.label}
            type="button"
            role="option"
            className="project-pages-color-swatch"
            title={color.label}
            aria-label={`${mode === "highlight" ? "Highlight" : "Text color"}: ${color.label}`}
            onMouseDown={(event) => {
              event.preventDefault();
              onPick(color.value);
            }}
          >
            <span
              className="project-pages-color-chip"
              style={mode === "highlight" ? { background: color.value || "transparent" } : { color: color.value || "inherit" }}
            >
              {color.value ? "A" : "×"}
            </span>
          </button>
        ))}
      </div>
    </div>
  );
}

function normalizeHref(url) {
  const raw = String(url || "").trim();
  if (/^(https?:|mailto:|tel:|file:|ftp:)/i.test(raw)) return raw;
  if (/^[\w.+-]+@[\w.-]+\.\w+$/.test(raw)) return `mailto:${raw}`;
  if (/^\/\//.test(raw)) return `https:${raw}`;
  return `https://${raw}`;
}

function mount(container, options = {}) {
  if (!container) return false;
  currentOptions = options || {};
  if (root && mountedContainer !== container) {
    root.unmount();
    root = null;
  }
  mountedContainer = container;
  if (!root) root = createRoot(container);
  root.render(<PageEditorApp />);
  return true;
}

function unmount() {
  if (root) root.unmount();
  root = null;
  mountedContainer = null;
  currentContext = null;
  flushCurrentEditor = null;
}

function setDocument(pageContext) {
  currentContext = pageContext || null;
  emitContext();
}

async function flushSave() {
  if (flushCurrentEditor) await flushCurrentEditor();
}

const ProjectPagesEditorApi = { flushSave, mount, setDocument, unmount };

if (typeof window !== "undefined") {
  window.ProjectPagesEditor = ProjectPagesEditorApi;
}

export { flushSave, mount, setDocument, unmount };
