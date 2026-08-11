import {
  CvssVector,
  baseScore,
  calculatorUrl,
  formatVector,
  isVectorLine,
  severityFor,
} from "./cvss";
import { isRiskHeading } from "./severity";

/* global Word */

export interface PreviousVector {
  text: string;
  hyperlink: string;
}

export interface RiskUndo {
  bookmark: string;
  previousRisk: string;
  /** Present when an existing vector line was overwritten rather than a new one added. */
  previousVector?: PreviousVector;
}

export interface RiskWritten {
  /** What the Risk line now reads after "Risk: ". */
  rating: string;
  vector: string;
  undo: RiskUndo;
}

/**
 * Reused for every insertion: only the most recent one is undoable, and
 * `insertBookmark` drops an existing bookmark of the same name first, so these do not
 * accumulate.
 */
const RISK_BOOKMARK = "_ptrisk";

/**
 * Write a CVSS rating onto the Risk line the cursor is on, with the vector beneath it.
 *
 * The Risk line is rewritten whole, which is what guarantees the exact
 * `Risk: <Severity> (<score>)` form that `severity.ts` will accept — the rating this
 * writes has to be one the sorting and table features can read back.
 *
 * No OOXML is involved: replacing paragraph text keeps each paragraph's own style, and
 * `Range.hyperlink` links the vector without needing a relationship part.
 */
export async function insertRisk(vector: CvssVector): Promise<RiskWritten> {
  return Word.run(async (context) => {
    const risk = context.document.getSelection().paragraphs.getFirst();
    risk.load("text");
    const below = risk.getNextOrNullObject();
    below.load("text");
    await context.sync();

    if (!isRiskHeading(risk.text)) {
      throw new Error(
        `Put the cursor on a "Risk:" line first — this one reads "${risk.text.trim() || "(empty)"}".`
      );
    }

    // Overwrite a vector already sitting under the Risk line rather than stacking another.
    const replacing = !below.isNullObject && isVectorLine(below.text);
    let previousVector: PreviousVector | undefined;
    if (replacing) {
      const existing = below.getRange("Whole");
      existing.load("hyperlink");
      await context.sync();
      previousVector = { text: below.text, hyperlink: existing.hyperlink ?? "" };
    }

    const score = baseScore(vector);
    const rating = `${severityFor(score)} (${score.toFixed(1)})`;
    const vectorText = formatVector(vector);
    const previousRisk = risk.text;

    risk.insertText(`Risk: ${rating}`, Word.InsertLocation.replace);

    const line = replacing ? below : risk.insertParagraph(vectorText, Word.InsertLocation.after);
    if (replacing) {
      line.insertText(vectorText, Word.InsertLocation.replace);
    }
    line.styleBuiltIn = "Normal";
    line.getRange("Whole").hyperlink = calculatorUrl(vector);

    risk.getRange("Whole").insertBookmark(RISK_BOOKMARK);
    await context.sync();

    return {
      rating,
      vector: vectorText,
      undo: { bookmark: RISK_BOOKMARK, previousRisk, previousVector },
    };
  });
}

/** Put the Risk line, and the vector below it, back as they were. */
export async function undoRisk(undo: RiskUndo): Promise<void> {
  return Word.run(async (context) => {
    const marked = context.document.getBookmarkRangeOrNullObject(undo.bookmark);
    marked.load("text");
    await context.sync();

    if (marked.isNullObject) {
      throw new Error("That risk rating is no longer where it was written.");
    }

    const risk = marked.paragraphs.getFirst();
    const below = risk.getNextOrNullObject();
    below.load("text");
    await context.sync();

    risk.insertText(undo.previousRisk, Word.InsertLocation.replace);

    if (!below.isNullObject && isVectorLine(below.text)) {
      if (undo.previousVector) {
        below.insertText(undo.previousVector.text, Word.InsertLocation.replace);
        if (undo.previousVector.hyperlink) {
          below.getRange("Whole").hyperlink = undo.previousVector.hyperlink;
        }
      } else {
        below.delete();
      }
    }

    context.document.deleteBookmark(undo.bookmark);
    await context.sync();
  });
}
