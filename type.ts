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
