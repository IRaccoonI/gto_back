// function docx_replace_regex(doc_obj, regex, getVal, style):
//     def subLetters(matcobj):
//         return getVal(matcobj[0][1:-1])

import { round } from "./math";
import { Est, EstRes, EstTable, NOT_DEFINE } from "../const/excel";

//     tt = [pp.text for pp in doc_obj.paragraphs]
//     pass
//     for p in doc_obj.paragraphs:
//         if re.search(regex, p.text):
//             p.text = re.sub(regex, subLetters, p.text)
//             checl_paragraph4link(p)
//             try:
//                 p.style = doc_obj.styles['Normal']
//             except:
//                 pass

//     for table in doc_obj.tables:
//         for row in table.rows:
//             try:
//                 row.cells
//             except:
//                 break
//             for cell in row.cells:
//                 docx_replace_regex(cell, regex, getVal, style)

// LINK_REGEX = r"https?:\/\/.*"

// def checl_paragraph4link(p):
//     if re.search(LINK_REGEX, p.text):
//         link_re = re.search(LINK_REGEX, p.text)
//         link = p.text[link_re.regs[0][0]:link_re.regs[0][1]]
//         p.text = p.text[:link_re.regs[0][0]]

//         add_hyperlink(p, link, link)

// def add_hyperlink(paragraph, text, url):
//     # This gets access to the document.xml.rels file and gets a new relation id value
//     part = paragraph.part
//     r_id = part.relate_to(
//         url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

//     # Create the w:hyperlink tag and add needed values
//     hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
//     hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )

//     # Create a w:r element and a new w:rPr element
//     new_run = docx.oxml.shared.OxmlElement('w:r')
//     rPr = docx.oxml.shared.OxmlElement('w:rPr')

//     # Join all the xml elements together add add the required text to the w:r element
//     new_run.append(rPr)
//     new_run.text = text
//     hyperlink.append(new_run)

//     # Create a new Run object and add the hyperlink into it
//     r = paragraph.add_run()
//     r._r.append(hyperlink)

//     # A workaround for the lack of a hyperlink style (doesn't go purple after using the link)
//     # Delete this if using a template that has the hyperlink style in it
//     r.font.color.theme_color = MSO_THEME_COLOR_INDEX.HYPERLINK
//     r.font.underline = True

//     return hyperlink

export function letterToInd(letter: string): number {
  let res = -1;
  [...letter].forEach((ch) => {
    res = (res + 1) * 26 + ch.charCodeAt(0) - 65;
  });
  return res;
}

const defaultEstTable = {
  3: ["Ниже среднего", "Средний", "Выше среднего"],
  5: ["Низкая", "Ниже среднего", "Средний", "Выше среднего", "Высокая"],
};

// TStratifyEstimation = Dict[str, Dict[int, List[float]]]

export function getEstimation(
  sex: string,
  age: number,
  value: number,
  stratifyEstimation: Est,
  estTable?: EstTable
): EstRes {
  const neededEstList = stratifyEstimation[sex][age];
  const isRise = neededEstList[0] < neededEstList[1];

  const estRes = neededEstList.reduce(
    (count, est) =>
      count + Number((!isRise && value < est) || (isRise && value > est)),
    0
  );

  estTable = estTable || defaultEstTable;
  const estStr = estTable[neededEstList.length + 1][estRes];
  const estSplit = `${estRes + 1}/${neededEstList.length + 1}`;

  return [estRes + 1, estStr, estSplit];
}

export function getConclusion(
  val: number | NOT_DEFINE,
  conclusions: number[],
  strs: string[]
) {
  if (val === NOT_DEFINE) {
    return NOT_DEFINE;
  }
  let res = 0;
  while (res < conclusions.length && val > conclusions[res]) {
    res += 1;
  }
  return strs[res];
}

function getDate(d: string) {
  if (/^d{1,2}.d{1,2}.d{4}$/.test(d)) {
    const [day, month, year] = d.split(".");
    return new Date(+year, +month, +day);
  }
  // if (/^d{4}-d{1,2}-d{1,2} d{2}:d{2}:d{2}$/.test(d)) {
  return new Date(d);
}

export function dateDiffInYear(
  _d1: string | Date,
  _d2: string | Date,
  floor: boolean = false
): number {
  const d1 = typeof _d1 === "string" ? getDate(_d1) : _d1;
  const d2 = typeof _d2 === "string" ? getDate(_d2) : _d2;
  return (floor ? Math.floor : round)(
    Math.abs(
      d2.getMonth() +
        d2.getFullYear() * 12 -
        d1.getMonth() -
        d1.getFullYear() * 12
    ) / 12
  );
}
