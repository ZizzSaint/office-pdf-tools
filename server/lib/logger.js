/** 极简日志。 */
const LEVELS = { error: 0, warn: 1, info: 2, debug: 3 };
const level = LEVELS[process.env.LOG_LEVEL] ?? LEVELS.info;

function stamp() {
  return new Date().toISOString().slice(11, 19);
}

function emit(kind, args) {
  if (LEVELS[kind] > level) return;
  const prefix = `[${stamp()}] ${kind.toUpperCase().padEnd(5)}`;
  (kind === "error" ? console.error : console.log)(prefix, ...args);
}

export const log = {
  error: (...a) => emit("error", a),
  warn: (...a) => emit("warn", a),
  info: (...a) => emit("info", a),
  debug: (...a) => emit("debug", a),
};
