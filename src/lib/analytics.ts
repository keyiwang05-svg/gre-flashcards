type TrackValue = string | number | boolean | null | undefined;
type TrackProps = Record<string, TrackValue>;

function cleanProps(props: TrackProps = {}) {
  return Object.fromEntries(Object.entries(props).filter(([, value]) => value !== undefined));
}

export function track(event: string, props: TrackProps = {}) {
  const payload = {
    event,
    props: cleanProps(props),
    path: typeof window !== "undefined" ? window.location.pathname : "",
    ts: Date.now(),
  };

  console.log("[track]", payload);
}

export function trackPageView(page = "home") {
  track("page_view_home", { page });
}
