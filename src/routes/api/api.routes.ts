import koaRouter from "koa-router";
import { demo } from "./api.controller";
import { validateParams } from "../../middleware/validate-params";

/**
 * A simple module to demonstrate declarative parameter validation.
 */
export const apiRouter = new koaRouter().post(
  "/foo-is-required",
  validateParams<string>(["query"], ["foo"]),
  demo
);
