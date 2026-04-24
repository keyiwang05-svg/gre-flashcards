import posthog from "posthog-js";

type TrackProps = Record<string, string | number | boolean | null | undefined>;

let posthogReady = false;

const BLOCKED_PROP_KEYS = new Set([
  "word",
  "prompt_word",
  "selected_choice",
  "correct_answer",
  "pair_a",
  "pair_b",
  "query",
  "search",
]);

function cleanProps(props: TrackProps = {}) {
  return Object.fromEntries(
    Object.entries(props).filter(([key, value]) => {
      if (BLOCKED_PROP_KEYS.has(key)) return false;
      return value !== undefined && value !== "";
    })
  );
}

export function initProductAnalytics() {
  if (posthogReady) return;
  if (typeof window === "undefined") return;

  const key = import.meta.env.VITE_POSTHOG_KEY;
  const host = import.meta.env.VITE_POSTHOG_HOST;

  if (!key || !host) {
    console.warn("[posthog] Missing VITE_POSTHOG_KEY or VITE_POSTHOG_HOST");
    return;
  }

  posthog.init(key, {
    api_host: host,
    autocapture: false,
    capture_pageview: false,
    capture_pageleave: false,
    disable_session_recording: true,
    disable_surveys: true,
  });

  posthogReady = true;
}

export function track(event: string, props: TrackProps = {}) {
  const payload = {
    ...cleanProps(props),
    app_name: "gre_flashcards",
    app_env: import.meta.env.MODE,
    hostname: typeof window !== "undefined" ? window.location.hostname : "",
    path: typeof window !== "undefined" ? window.location.pathname : "",
  };

  console.log("[track]", event, payload);

  if (!posthogReady) {
    initProductAnalytics();
  }

  if (posthogReady) {
    posthog.capture(event, payload);
  }
}

export function trackPageView(page = "home") {
  track("page_view_home", { page });
}