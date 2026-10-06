/** 统一错误类型与 HTTP 映射。 */
export class AppError extends Error {
  /**
   * @param {string} message 面向用户的中文说明
   * @param {{code?:string, status?:number, hint?:string, cause?:Error, detail?:any}} [opts]
   */
  constructor(message, opts = {}) {
    super(message);
    this.name = "AppError";
    this.code = opts.code || "internal-error";
    this.status = opts.status || 500;
    this.hint = opts.hint;
    this.detail = opts.detail;
    if (opts.cause) this.cause = opts.cause;
  }

  toJSON() {
    return { code: this.code, message: this.message, hint: this.hint, detail: this.detail };
  }
}

export const badRequest = (message, opts = {}) => new AppError(message, { status: 400, code: "bad-request", ...opts });
export const notFound = (message, opts = {}) => new AppError(message, { status: 404, code: "not-found", ...opts });
export const unsupported = (message, opts = {}) => new AppError(message, { status: 415, code: "unsupported-format", ...opts });
export const engineMissing = (message, opts = {}) => new AppError(message, { status: 503, code: "engine-unavailable", ...opts });
export const timeoutError = (message, opts = {}) => new AppError(message, { status: 504, code: "timeout", ...opts });

export function toAppError(err) {
  if (err instanceof AppError) return err;
  return new AppError(err?.message || String(err), { code: "internal-error", status: 500, cause: err });
}

/** 包装 async 路由处理器。 */
export function wrap(handler) {
  return (req, res, next) => Promise.resolve(handler(req, res, next)).catch(next);
}
