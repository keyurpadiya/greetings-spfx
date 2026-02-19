export interface IGreetingsProps {
  description: string;
  greetingText: string;
  showFirstNameOnly: boolean;
  fontSize: string;
  fontSizeUnit?: string;
  fontSizeValue?: string;
  fontStyle: string;
  backgroundImageUrl: string | IFilePickerResult | null; // Update this line
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context:any;
   fontColor: string; // ✅ NEW: Manual font color
  textAlignment: string; // ✅ NEW: Text alignment
  
}
export interface IFilePickerResult {
  fileAbsoluteUrl?: string;
  spItemUrl?: string;
  fileName?: string;
  [key: string]: any;
}