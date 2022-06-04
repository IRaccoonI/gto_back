import { Context } from "koa";
import md5File from "md5-file";
import convertExcelToJson from "convert-excel-to-json";
import { isDate } from "../../lib/utils";
import { parseExcel } from "../../lib/parseExcel";

export async function uploadGto(ctx: Context) {
  const file = ctx.request.files.file as never as File & { path: string };
  console.log(md5File.sync(file.path));
  const excel = convertExcelToJson({ sourceFile: file.path });
  const res: Record<string, string | number>[] = Object.keys(excel)
    .map((className) =>
      excel[className].reduce(
        (res, row, idx) => (idx < 5 ? res : [...res, { className, ...row }]),
        []
      )
    )
    .reduce((res, rows) => [...res, ...rows], [])
    .map((row: Record<string, string | number | Date>) =>
      Object.assign(
        {},
        ...Object.keys(row).map((key) => {
          const val = row[key];
          if (isDate(val)) {
            return { [key]: val.toISOString() };
          }
          return { [key]: val };
        })
      )
    );

  ctx.response.body = res.map(parseExcel);
}

/**
 * Demo Error Responder: Deliberataly return 500 error for testing.
 */
export async function error(ctx: Context) {
  ctx.status = 500;
  ctx.message = "App Error (this is intentional)!";
}

/**
 * Demo Error Responder: Deliberataly return 500 error without message for testing.
 */
export async function errorWithoutMessage() {
  // eslint-disable-next-line no-console
  console.log("About to throw an error deliberately, ignore it.");
  throw new Error("");
}
