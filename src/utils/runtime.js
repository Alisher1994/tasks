export function isHostedRuntime() {
  return typeof window !== "undefined"
    && (window.location.protocol === "http:" || window.location.protocol === "https:");
}
