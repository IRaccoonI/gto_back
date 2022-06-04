import {
  dateDiffInYear,
  getConclusion,
  getEstimation,
  letterToInd,
} from "./utilsExcel";
import * as excelConst from "../const/excel";
import { round } from "./math";
import tmpConst from "./tmpConst";
import { F } from "ramda";

interface Recommendations {
  speed: boolean;
  flex: boolean;
  strength: boolean;
  coordination: boolean;
}

export function parseExcel(
  row: Record<string, string | number>
): Record<string, string | number | Date> {
  const rec: Recommendations = {
    coordination: false,
    flex: false,
    speed: false,
    strength: false,
  };

  let addVals: Record<string, string>;

  prepareBefore(row);
  initAddVals();
  initConclusion();
  initRecommendations();

  return getRes();

  // Private
  function getRes() {
    return Object.assign(
      {},
      ...Object.keys(addVals).map((k) => {
        let val: string = addVals[k];
        if (val === "") {
          return { [k]: val };
        }
        if (!isNaN(+val)) {
          return { [k]: round(+val, 2) };
        }
        if (!isNaN(new Date(val).getMonth())) {
          return { [k]: new Date(val) };
        }
        return { [k]: val };
      })
    );
  }

  function prepareBefore(row: Record<string, string | number>) {
    addVals = Object.assign(
      {},
      ...Object.keys(row).map((k) => {
        const val = `${row[k]}`.replace(/\s+/g, " ").trim();
        if (k === "E") {
          return { [k]: val[0].toUpperCase() };
        }
        return { [k]: val };
      })
    );
  }

  function initAddVals() {
    let age = getAge();
    if (age < 7) {
      age = 7;
    }
    if (age > 12) {
      age = 12;
    }
    let sex = getVal("E")[0].toLowerCase();

    function getEst(
      val: string | number,
      est: excelConst.Est,
      estTable?: excelConst.EstTable
    ): excelConst.EstRes {
      try {
        const fVal = Number(val);
        return getEstimation(sex, age, fVal, est, estTable);
      } catch {
        return [excelConst.NOT_DEFINE, "", ""];
      }
    }

    addVals["AGE"] = "" + getFactAge();

    let pullUps;
    try {
      pullUps = sex == "ж" ? getVal("AS") : getVal("AT");
      addVals["PullUps"] = pullUps;
    } catch {}

    let strs: excelConst.EstTable = {
      5: ["Низкий", "Ниже среднего", "Средний", "Выше среднего", "Высокий"],
    };
    try {
      const bodyMassIndex =
        Math.round(
          (Number(getVal("Q")) / (Number(getVal("P")) / 100) ** 2) * 10
        ) / 10;
      addVals["BodyMassIndex"] = "" + bodyMassIndex;

      addToVal("BL_EST", getEst(getVal("P"), excelConst.BL_EST));
      addToVal("BWI_EST", getEst(bodyMassIndex, excelConst.BWI_EST, strs));
    } catch {}

    addToVal("BW_EST", getEst(getVal("Q"), excelConst.BW_EST, strs));

    let uRounded: number | excelConst.NOT_DEFINE = null;
    if (getVal("U") === excelConst.NOT_DEFINE) {
      uRounded = excelConst.NOT_DEFINE;
    } else {
      uRounded = round(round(Number(getVal("U")), 1));
    }
    addVals["U_ROUNDED"] = "" + uRounded;

    addToVal("LVC_EST", getEst(getVal("U"), excelConst.LVC_EST));
    addToVal("CS_EST", getEst(getVal("R"), excelConst.CS_EST));
    addToVal("LPS_EST", getEst(getVal("W"), excelConst.LPS_EST));
    addToVal("RPS_EST", getEst(getVal("V"), excelConst.RPS_EST));
    if (
      uRounded === excelConst.NOT_DEFINE ||
      getVal("Q") == excelConst.NOT_DEFINE
    ) {
      addVals["LI"] = excelConst.NOT_DEFINE;
    } else {
      addVals["LI"] = "" + round(uRounded / +getVal("Q"), 1);
    }

    addToVal("LI_EST", getEst(getVal("LI"), excelConst.LI_EST));
    strs = {
      3: ["Ниже нормы", "Норма", "Выше нормы"],
    };

    addToVal("HR_EST", getEst(getVal("AC"), excelConst.HR_EST, strs));
    addToVal("SP_EST", getEst(getVal("AA"), excelConst.SP_EST, strs));
    addToVal("DP_EST", getEst(getVal("AB"), excelConst.DP_EST, strs));

    strs = {
      5: ["Низкий", "Ниже среднего", "Средний", "Выше среднего", "Высокий"],
    };
    addToVal("FB_EST", getEst(getVal("AU"), excelConst.FB_EST, strs));
    addToVal("CR_3_10_EST", getEst(getVal("AQ"), excelConst.CR_3_10_EST, strs));
    addToVal("PU_EST", getEst(pullUps, excelConst.PU_EST, strs));
    addToVal("R_30_EST", getEst(getVal("AO"), excelConst.R_30_EST, strs));
    addToVal("LJ_EST", getEst(getVal("AR"), excelConst.LJ_EST, strs));
    addToVal("R_6M_EST", getEst(getVal("AP"), excelConst.R_6M_EST, strs));

    let PWC170_VAL;
    if (
      getVal("AM") === excelConst.NOT_DEFINE ||
      getVal("Q") === excelConst.NOT_DEFINE
    ) {
      PWC170_VAL = excelConst.NOT_DEFINE;
    } else {
      PWC170_VAL = "" + round(+getVal("AM") / +getVal("Q"), 1);
    }
    addVals["PWC170_V"] = "" + PWC170_VAL;
    addToVal("PWC170", getEst(PWC170_VAL, excelConst.PWC170_EST));

    const avgPoint = _getRecAvgPoint();
    const tmp_lst = [1.5, 2, 3, 4, 999];
    const strs1 = [
      "Низкий",
      "Ниже среднего",
      "Среднее",
      "Выше среднего",
      "Высокий",
    ];
    if (avgPoint === excelConst.NOT_DEFINE) {
      addVals["PSCS"] = avgPoint;
      addToVal("PSCS", [excelConst.NOT_DEFINE, "", ""]);
    } else {
      addVals["PSCS"] = "" + round(avgPoint, 1);
      const i = tmp_lst.reduce(
        (count, val) => count + Number(val < avgPoint),
        0
      );
      const est: excelConst.EstRes = [i + 1, strs1[i], `${i + 1}/5`];
      addToVal("PSCS", est);
    }
  }

  function initConclusion() {
    function avg(vals: string[], ignoreND = false) {
      if (ignoreND) vals = vals.filter((val) => val !== excelConst.NOT_DEFINE);
      if (vals.some((val) => isNaN(+val))) return excelConst.NOT_DEFINE;

      const sum = vals.reduce((sum, val) => sum + +val, 0);
      const len = vals.length;

      return round(sum / len, 1);
    }

    function avgVals(letters: string[], ignoreND = false) {
      return avg(
        letters.map((letter) => getVal(letter)),
        ignoreND
      );
    }

    // Physical development
    const PD_CONC = [getVal("BL_EST_RES"), getVal("BW_EST_RES")];
    if (PD_CONC.every((val) => !isNaN(+val))) {
      addVals["PD_CONC"] = excelConst.DISHARMONIOUS_DEVELOPMENT_PAIRS.includes(
        PD_CONC.map((val) => +val)
      )
        ? "Дисгармоничное"
        : "Гармоничное";
    } else {
      addVals["PD_CONC"] = excelConst.NOT_DEFINE;
    }

    //  state of physiometric functions
    //  ((l + r) / 2 + lifeindex) / 2 + дальше в 1 пункте заключения тоже пересчитать
    let strs = [
      "Низкое",
      "Ниже среднего",
      "Среднее",
      "Выше среднего",
      "Высокое",
    ];
    let vals = [1.4, 2.4, 3.4, 4.4];
    const lr = avgVals(["LPS_EST_RES", "RPS_EST_RES"]);
    if (lr !== excelConst.NOT_DEFINE && !isNaN(+getVal("LI_EST_RES"))) {
      addVals["SPF_CONC"] = getConclusion(
        (lr + +getVal("LI_EST_RES")) / 2,
        vals,
        strs
      );
    } else {
      addVals["SPF_CONC"] = excelConst.NOT_DEFINE;
    }

    //  level of morphofunctional development
    strs = ["Низкий", "Ниже среднего", "Средний", "Выше среднего", "Высокий"];
    vals = [1.4, 2.4, 3.4, 4.4];
    let lmd: excelConst.NOT_DEFINE | number;

    lmd = avgVals(["LPS_EST_RES", "RPS_EST_RES"], true);

    if (lmd !== excelConst.NOT_DEFINE && !isNaN(+getVal("LI_EST_RES"))) {
      lmd = (lmd + +getVal("LI_EST_RES")) / 2;
    }

    if (lmd !== excelConst.NOT_DEFINE && !isNaN(+getVal("BWI_EST_RES"))) {
      lmd = (lmd + +getVal("BWI_EST_RES")) / 2;
    }

    addVals["LMD_CONC"] = getConclusion(lmd, vals, strs);

    //  State of indicators of central hemodynamics
    strs = ["Ниже нормы", "Норма", "Выше нормы"];
    vals = [1.4, 2.4];
    if (
      getVal("HR_EST_RES") === getVal("SP_EST_RES") &&
      getVal("SP_EST_RES") == getVal("DP_EST_RES")
    ) {
      addVals["SICH_CONC"] = getVal("HR_EST_STR");
    } else {
      addVals["SICH_CONC"] = excelConst.NOT_DEFINE;
    }

    //  General level of physical fitness
    strs = ["Низкий", "Ниже среднего", "Средний", "Выше среднего", "Высокий"];
    vals = [1.4, 2.4, 3.4, 4.4];
    addVals["GLPF_CONC"] = getConclusion(
      avgVals(
        [
          "R_6M_EST_RES",
          "LJ_EST_RES",
          "R_30_EST_RES",
          "PU_EST_RES",
          "CR_3_10_EST_RES",
          "FB_EST_RES",
        ],
        true
      ),
      vals,
      strs
    );
  }

  function initRecommendations() {
    addVals["RECOMMEND"] = _Recommendations();
    // end_rec_val =  addVals["CR_3_10_EST_RES"]
    // if end_rec_val == NOT_DEFINE:
    //      addVals["END_REC"] = NOT_DEFINE
    // else:
    //     // to_sum = [-5, 0, 0, 5, 10][end_rec_val - 1]
    const age = getFactAge();
    const limits = [round((220 - age) * 0.7), round((220 - age) * 0.8)];
    const val = excelConst.STAMINA_REC(limits);
    addVals["END_REC"] = val;

    if (!isNaN(+getVal("R_30_EST_RES")) && +getVal("R_30_EST_RES") <= 2) {
      rec.speed = true;
      addVals["SPEED_REC"] =
        "Для улучшения показателей скорости: " + excelConst.SPEED_REC_LINK;
    } else {
      addVals["SPEED_REC"] = "";
    }

    if (!isNaN(+getVal("FB_EST_RES")) && +getVal("FB_EST_RES") <= 2) {
      rec.flex = true;
      addVals["FLEX_REC"] =
        "Для улучшения показателей гибкости: " + excelConst.SPEED_REC_LINK;
    } else {
      addVals["FLEX_REC"] = "";
    }

    if (!isNaN(+getVal("LJ_EST_RES")) && +getVal("LJ_EST_RES") <= 2) {
      rec.coordination = true;
      addVals["COORDINATION_REC"] =
        "Для улучшения показателей координации: " +
        excelConst.COORDINATION_REC_LINK;
    } else {
      addVals["COORDINATION_REC"] = "";
    }

    if (!isNaN(+getVal("PU_EST_RES")) && +getVal("PU_EST_RES") <= 2) {
      rec.strength = true;
      addVals["STRENGTH_REC"] =
        "Для улучшения показателей силы: " + excelConst.STRENGTH_REC_LINK;
    } else {
      addVals["STRENGTH_REC"] = "";
    }
  }

  function _getRecAvgPoint() {
    // BC/10 + (100 - BH) / 10
    let BC_STR = getVal("BC");
    let BH_STR = getVal("BH");
    if (BC_STR === excelConst.NOT_DEFINE || BH_STR === excelConst.NOT_DEFINE)
      return excelConst.NOT_DEFINE;
    if (isNaN(+BC_STR) || isNaN(+BH_STR)) return excelConst.NOT_DEFINE;

    const BC = +BC_STR;
    const BH = +BH_STR;

    const point0 = BC / 10 + (100 - BH) / 10;

    // BO
    const BO_STR = getVal("BO");
    if (BO_STR == excelConst.NOT_DEFINE || isNaN(+BO_STR))
      return excelConst.NOT_DEFINE;
    const BO = +BO_STR;

    let tmp_lst = [1000, 2000, 4000, 4900];
    let point1 = 0;
    while (point1 < tmp_lst.length && BO >= tmp_lst[point1]) {
      point1 += 1;
    }
    point1 += 1;

    // ЕСЛИ(P2<=0,5;"5";ЕСЛИ(P2<=0,89;"4";ЕСЛИ(P2<=1,1;"3";ЕСЛИ(P2<=2;"2";"1")))))
    // CF .5, .89, 1, 2
    const CF_STR = getVal("CF");
    if (CF_STR === excelConst.NOT_DEFINE || isNaN(+CF_STR))
      return excelConst.NOT_DEFINE;
    const CF = +CF_STR;

    tmp_lst = [2, 1, 0.89, 0.5];
    let point2 = 0;
    while (point2 < tmp_lst.length && CF < tmp_lst[point2]) {
      point2 += 1;
    }
    point2 += 1;

    // ЕСЛИ(BN>100;ЕСЛИ(BV>240;"I ТИП";"II ТИП");
    // ЕСЛИ(BN>25;ЕСЛИ(BV>240;"III ТИП";"- ");
    // ЕСЛИ(BV>500;"IV ТИП";"  - ")))
    // BN BV
    const BN_STR = getVal("BN");
    const BV_STR = getVal("BV");
    if (BN_STR === excelConst.NOT_DEFINE || BV_STR === excelConst.NOT_DEFINE)
      return excelConst.NOT_DEFINE;
    if (isNaN(+BN_STR) || isNaN(+BN_STR)) return excelConst.NOT_DEFINE;
    const BN = +BN_STR;
    const BV = +BV_STR;

    let tmp_type = 2;
    if (BN > 100)
      if (BV > 240) tmp_type = 1;
      else tmp_type = 2;
    else if (BN > 25) tmp_type = 3;
    else tmp_type = 4;

    //  2 -> 1, 3 -> 5, * -> 2
    let point3 = 0;
    if (tmp_type === 2) point3 = 1;
    else if (tmp_type == 3) point3 = 5;
    else point3 = 2;

    return (point0 + point1 + point2 + point3) / 4;
  }

  function _Recommendations() {
    const avgPoint = _getRecAvgPoint();
    const tmp_lst = [1.5, 2, 3, 4, 5];
    if (avgPoint === excelConst.NOT_DEFINE) {
      return excelConst.NOT_DEFINE;
    }

    for (let i = 0; i < excelConst.RECOMMENDATIONS.length; i++) {
      if (avgPoint < tmp_lst[i]) {
        return excelConst.RECOMMENDATIONS[i];
      }
    }

    return excelConst.RECOMMENDATIONS.at(-1);
  }

  function addToVal(propName: string, est: excelConst.EstRes) {
    addVals[propName + "_RES"] = "" + est[0];
    addVals[propName + "_STR"] = est[1];
    addVals[propName + "_SPL"] = est[2];
    addVals[propName + "_STR_IMP"] =
      est[1] !== "" ? est[1] : excelConst.NOT_DEFINE;
  }

  function getVal(symbol: string): string {
    let res = excelConst.NOT_DEFINE;
    let lower = false;
    let notDefineOne = false;
    let notDefineTwo = false;
    let notDefineEmpty = false;
    if (symbol.includes("_LOWER")) {
      lower = true;
      symbol = symbol.replace("_LOWER", "");
    }
    if (symbol.includes("_ND")) {
      notDefineOne = true;
      symbol = symbol.replace("_ND", "");
    }
    if (symbol.includes("_ND1")) {
      notDefineTwo = true;
      symbol = symbol.replace("_ND1", "");
    }
    if (symbol.includes("_EMPTYND")) {
      notDefineEmpty = true;
      symbol = symbol.replace("_EMPTYND", "");
    }

    if (addVals[symbol]) {
      res = addVals[symbol];
    } else {
      const findInd = letterToInd(symbol);
      try {
        res = "" + row[findInd];
      } catch {
        res = excelConst.NOT_DEFINE + ".";
      }

      if (res == "nan") {
        res = excelConst.NOT_DEFINE;
      }
    }

    if (res == excelConst.NOT_DEFINE && notDefineOne) {
      res = excelConst.NOT_DEFINE_ONE;
    }
    if (res == excelConst.NOT_DEFINE && notDefineTwo) {
      res = excelConst.NOT_DEFINE_TWO;
    }
    if (res == excelConst.NOT_DEFINE && notDefineEmpty) {
      res = "";
    }
    if (lower) {
      res = `${res}`.toLowerCase();
    }
    return res;
  }

  function getAge() {
    const _d1 = getVal("H");
    const _d2 = getVal("G");
    return dateDiffInYear(_d1, _d2);
  }

  function getFactAge() {
    const _d1 = getVal("H");
    const _d2 = new Date();
    return dateDiffInYear(_d1, _d2, true);
  }

  // function save(templateNames, saveDirName, recFiles){

  //     if  isError:
  //         return

  //     school =  getVal("D")

  //     try:
  //         if not os.path.isdir(OUT_DIR):
  //             os.mkdir(OUT_DIR)
  //         if not os.path.isdir(f'{OUT_DIR}/{saveDirName}'):
  //             os.mkdir(f'{OUT_DIR}/{saveDirName}')
  //         if not os.path.isdir(f'{OUT_DIR}/{saveDirName}/{school}'):
  //             os.mkdir(f'{OUT_DIR}/{saveDirName}/{school}')
  //         if not os.path.isdir(f'{OUT_DIR}/{saveDirName}/{school}/{ className}'):
  //             os.mkdir(f'{OUT_DIR}/{saveDirName}/{school}/{ className}')
  //         if not os.path.isdir(f'{OUT_DIR}/{saveDirName}/{school}/{ className}/{ generalInformation.fullName}'):
  //             os.mkdir(
  //                 f'{OUT_DIR}/{saveDirName}/{school}/{ className}/{ generalInformation.fullName}')
  //     except:
  //         print( getVal( generalInformation.fullName),
  //               "Что-то вообще не так с записью")

  //     FINISH_OUT_DIR = f'{OUT_DIR}/{saveDirName}/{school}/{ className}/{ generalInformation.fullName}'

  //     for templateName in templateNames[:1]:
  //         nPath = f'{FINISH_OUT_DIR}/{ generalInformation.fullName}.docx'
  //         if os.path.isfile(nPath):
  //             continue

  //         doc = Document(f'./{templateName}')

  //         regex1 = re.compile(r"%\w+%")

  //         style = doc.styles['Normal']
  //         font = style.font
  //         font.name = 'Times New Roman'
  //         font.size = Pt(12)

  //         docx_replace_regex(doc, regex1, lambda s:  getVal(s), style)
  //         if os.path.isfile(nPath):
  //             os.remove(nPath)

  //         try:
  //             doc.save(nPath)
  //             if  rec.flex:
  //                 shutil.copyfile(recFiles['flex'],
  //                                 f"{FINISH_OUT_DIR}/{recFiles['flex']}")
  //             if  rec.coordination:
  //                 shutil.copyfile(
  //                     recFiles['coordination'], f"{FINISH_OUT_DIR}/{recFiles['coordination']}")
  //             if  rec.speed:
  //                 shutil.copyfile(recFiles['speed'],
  //                                 f"{FINISH_OUT_DIR}/{recFiles['speed']}")
  //             if  rec.strength:
  //                 shutil.copyfile(
  //                     recFiles['strength'], f"{FINISH_OUT_DIR}/{recFiles['strength']}")
  //         except:
  //             print( getVal( generalInformation.fullName),
  //                   "Что-то вообще не так с записью")
  //                 }
}
