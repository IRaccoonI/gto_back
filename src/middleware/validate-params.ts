import { Context } from "koa";
import R from "ramda";
import { isNotEmptyObject, trusty } from "../lib/fp";
import { logger } from "../services/logger";

export interface Validator<T> {
  (val: T): boolean;
}

interface Container<T> {
  [param: string]: T;
}

/**
 * Middleware that checks for required parameters.
 *
 * @param { string[] } containerPath - Where the parameters live in the ctx
 * instance (session,[request, body], etc.).
 *
 * @param { string[] } params - the names of the params to check.
 *
 * @param { function } validator (optional) - a function to validate the
 * parameters in question. If this is omitted, a simple presence check will
 * be performed.
 */
export const validateParams =
  <T>(containerPath: string[], params: string[], validator?: Validator<T>) =>
  async (ctx: Context, next: Function) => {
    const container: Container<T> = R.path(containerPath, ctx);

    if (!container) {
      logger.warn("Invalid param container %j: %j", container, {
        requestId: ctx.requestId,
      });
      ctx.throw(400, "Bad request");
    }

    const validated = R.map(assertValid(container, validator), params);
    console.log(validated.filter(trusty));

    // if (!validated.every(isNotEmptyObject)) {
    if (validated.filter(trusty).length) {
      const required = validated
        .map((val) => val?.requiredError)
        .filter(trusty);
      const invalided = validated
        .map((val) => val?.invalidError)
        .filter(trusty);

      const requiredMessage = `${required.join(", ")} is required.${
        invalided.length ? "" : " "
      }`;
      const invalidedMessage = `${invalided.join(", ")} is invalid.`;

      ctx.throw(
        400,
        (required.length ? requiredMessage : "") +
          (invalided.length ? invalidedMessage : "")
      );
    }

    await next();
  };

const assertValid =
  <T>(container: Container<T>, validator?: Validator<T>) =>
  (param: string): { requiredError?: string; invalidError?: string } | null => {
    const result: { requiredError?: string; invalidError?: string } = {};
    if (!container[param]) {
      result.requiredError = param;
    }

    if (validator && !validator(container[param])) {
      result.invalidError = param;
    }

    return isNotEmptyObject(result) ? result : null;
  };
