import { JWT_SECRET } from "./../env/local";
import Koa from "koa";
import koaRouter from "koa-router";
import bodyParser from "koa-body";
import helmet from "koa-helmet";
import jwt from "koa-jwt";
import { logger } from "./services/logger";
import { errorResponder } from "./middleware/error-responder";
import { k } from "./project-env";
import { rootRouter } from "./routes/root.routes";
import { apiRouter } from "./routes/api/api.routes";
import koaBody from "koa-body";
import koaStatic from "koa-static";
import path from "path";

export const app = new Koa();

// Entry point for all modules.
const api = new koaRouter()
  .use("/", rootRouter.routes())
  .use("/api", apiRouter.routes());
// koa-body Intermediate Plug-in File Submission and form-data

// Configuring Static Resource Loading Middleware
// app.use(
//   koaStatic(
//     path.join("out") //Read static file directories
//   )
// );

app.use(api.routes()).use(api.allowedMethods());
function startFunction() {
  const PORT = process.env.PORT || 3000;
  logger.info(`Starting server on http://localhost:${PORT}`);
  app.listen(PORT);
}

/* istanbul ignore if */
if (require.main === module) {
  if (process.env.PROJECT_ENV === "staging") {
    const throng = require("throng");
    throng(startFunction);
  } else {
    startFunction();
  }
}
