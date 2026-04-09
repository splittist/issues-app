import { Criteria } from "./types";
import { dateToday } from "./utils";

export interface CriteriaOption {
  key: keyof Criteria;
  label: string;
}

export const defaultCriteria: Criteria = {
  redline: true,
  highlight: false,
  squareBrackets: false,
  comments: false,
  footnotes: false,
  endnotes: false,
};

export const criteriaOptions: CriteriaOption[] = [
  { key: 'redline', label: 'Includes redlining' },
  { key: 'highlight', label: 'Includes highlighted text' },
  { key: 'squareBrackets', label: 'Includes square brackets' },
  { key: 'comments', label: 'Includes comments' },
  { key: 'footnotes', label: 'Includes footnotes' },
  { key: 'endnotes', label: 'Includes endnotes' },
];

export const getDefaultOutputFileName = (): string => `report_${dateToday()}.docx`;
