import koaRouter from "koa-router";
import { uploadGto } from "./api.controller";
import koaBody from "koa-body";

/**
 * A simple module to demonstrate declarative parameter validation.
 */
export const apiRouter = new koaRouter().post(
  "/uploadGto",
  koaBody({
    formLimit: "1mb",
    multipart: true, // Allow multiple files to be uploaded
    formidable: {
      maxFileSize: 200 * 1024 * 1024, //Upload file size
      keepExtensions: true, //  Extensions to save images
      uploadDir: "uploads",
    },
  }),
  uploadGto
);
