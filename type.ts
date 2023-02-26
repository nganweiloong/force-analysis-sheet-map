export type SheetItem = {
  Story: string;
  Beam: string;
  UniqueName: string;
  "Output Case": string;
  "Case Type": string;
  "Step Type": string;
  Station: number;
  P: number;
  V2: number;
  V3: number;
  T: number;
  M2: number;
  M3: number;
  Element: string;
  "Element Station": number;
  Location: string;
};

export type NormalizedSheetItem = {
  [key: string]: SheetItem;
};

type SourceString =
  | "P source"
  | "V2 source"
  | "V3 source"
  | "T source"
  | "M2 source"
  | "M3 source"
  | "Row source";

export type SheetItemWithSources = SheetItem & Record<SourceString, string>;
