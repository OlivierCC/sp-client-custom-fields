/**
 * @file PropertyFieldIconPickerHost.tsx
 * Renders the controls for PropertyFieldIconPicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldIconPickerPropsInternal } from './PropertyFieldIconPicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
import GuidHelper from './GuidHelper';

/**
 * @interface
 * PropertyFieldIconPickerHost properties interface
 *
 */
export interface IPropertyFieldIconPickerHostProps extends IPropertyFieldIconPickerPropsInternal {
}

/**
 * @interface
 * PropertyFieldIconPickerHost state interface
 *
 */
export interface IPropertyFieldIconPickerHostState {
  isOpen: boolean;
  isHoverDropdown?: boolean;
  hoverFont?: string;
  selectedFont?: string;
  safeSelectedFont?: string;
  errorMessage?: string;
}

/**
 * @interface
 * Define a safe font object
 *
 */
interface ISafeFont {
  Name: string;
  SafeValue: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldIconPicker component
 */
export default class PropertyFieldIconPickerHost extends React.Component<IPropertyFieldIconPickerHostProps, IPropertyFieldIconPickerHostState> {
  /**
   * @var
   * Defines the font series
   */
  private fonts: ISafeFont[] = [
    {Name: "DecreaseIndentLegacy", SafeValue: 'ms-Icon--DecreaseIndentLegacy'},
    {Name: "IncreaseIndentLegacy", SafeValue: 'ms-Icon--IncreaseIndentLegacy'},
    {Name: "GlobalNavButton", SafeValue: 'ms-Icon--GlobalNavButton'},
    {Name: "InternetSharing", SafeValue: 'ms-Icon--InternetSharing'},
    {Name: "Brightness", SafeValue: 'ms-Icon--Brightness'},
    {Name: "MapPin", SafeValue: 'ms-Icon--MapPin'},
    {Name: "Airplane", SafeValue: 'ms-Icon--Airplane'},
    {Name: "Tablet", SafeValue: 'ms-Icon--Tablet'},
    {Name: "QuickNote", SafeValue: 'ms-Icon--QuickNote'},
    {Name: "ChevronDown", SafeValue: 'ms-Icon--ChevronDown'},
    {Name: "ChevronUp", SafeValue: 'ms-Icon--ChevronUp'},
    {Name: "Edit", SafeValue: 'ms-Icon--Edit'},
    {Name: "Add", SafeValue: 'ms-Icon--Add'},
    {Name: "Cancel", SafeValue: 'ms-Icon--Cancel'},
    {Name: "More", SafeValue: 'ms-Icon--More'},
    {Name: "Settings", SafeValue: 'ms-Icon--Settings'},
    {Name: "Video", SafeValue: 'ms-Icon--Video'},
    {Name: "Mail", SafeValue: 'ms-Icon--Mail'},
    {Name: "People", SafeValue: 'ms-Icon--People'},
    {Name: "Phone", SafeValue: 'ms-Icon--Phone'},
    {Name: "Pin", SafeValue: 'ms-Icon--Pin'},
    {Name: "Shop", SafeValue: 'ms-Icon--Shop'},
    {Name: "Link", SafeValue: 'ms-Icon--Link'},
    {Name: "Filter", SafeValue: 'ms-Icon--Filter'},
    {Name: "Zoom", SafeValue: 'ms-Icon--Zoom'},
    {Name: "ZoomOut", SafeValue: 'ms-Icon--ZoomOut'},
    {Name: "Microphone", SafeValue: 'ms-Icon--Microphone'},
    {Name: "Search", SafeValue: 'ms-Icon--Search'},
    {Name: "Camera", SafeValue: 'ms-Icon--Camera'},
    {Name: "Attach", SafeValue: 'ms-Icon--Attach'},
    {Name: "Send", SafeValue: 'ms-Icon--Send'},
    {Name: "FavoriteList", SafeValue: 'ms-Icon--FavoriteList'},
    {Name: "PageSolid", SafeValue: 'ms-Icon--PageSolid'},
    {Name: "Forward", SafeValue: 'ms-Icon--Forward'},
    {Name: "Back", SafeValue: 'ms-Icon--Back'},
    {Name: "Refresh", SafeValue: 'ms-Icon--Refresh'},
    {Name: "Share", SafeValue: 'ms-Icon--Share'},
    {Name: "Lock", SafeValue: 'ms-Icon--Lock'},
    {Name: "EMI", SafeValue: 'ms-Icon--EMI'},
    {Name: "MiniLink", SafeValue: 'ms-Icon--MiniLink'},
    {Name: "Blocked", SafeValue: 'ms-Icon--Blocked'},
    {Name: "FavoriteStar", SafeValue: 'ms-Icon--FavoriteStar'},
    {Name: "FavoriteStarFill", SafeValue: 'ms-Icon--FavoriteStarFill'},
    {Name: "ReadingMode", SafeValue: 'ms-Icon--ReadingMode'},
    {Name: "Remove", SafeValue: 'ms-Icon--Remove'},
    {Name: "Checkbox", SafeValue: 'ms-Icon--Checkbox'},
    {Name: "CheckboxComposite", SafeValue: 'ms-Icon--CheckboxComposite'},
    {Name: "CheckboxIndeterminate", SafeValue: 'ms-Icon--CheckboxIndeterminate'},
    {Name: "CheckMark", SafeValue: 'ms-Icon--CheckMark'},
    {Name: "BackToWindow", SafeValue: 'ms-Icon--BackToWindow'},
    {Name: "FullScreen", SafeValue: 'ms-Icon--FullScreen'},
    {Name: "Print", SafeValue: 'ms-Icon--Print'},
    {Name: "Up", SafeValue: 'ms-Icon--Up'},
    {Name: "Down", SafeValue: 'ms-Icon--Down'},
    {Name: "Delete", SafeValue: 'ms-Icon--Delete'},
    {Name: "Save", SafeValue: 'ms-Icon--Save'},
    {Name: "SIPMove", SafeValue: 'ms-Icon--SIPMove'},
    {Name: "EraseTool", SafeValue: 'ms-Icon--EraseTool'},
    {Name: "GripperTool", SafeValue: 'ms-Icon--GripperTool'},
    {Name: "Dialpad", SafeValue: 'ms-Icon--Dialpad'},
    {Name: "PageLeft", SafeValue: 'ms-Icon--PageLeft'},
    {Name: "PageRight", SafeValue: 'ms-Icon--PageRight'},
    {Name: "MultiSelect", SafeValue: 'ms-Icon--MultiSelect'},
    {Name: "Play", SafeValue: 'ms-Icon--Play'},
    {Name: "Pause", SafeValue: 'ms-Icon--Pause'},
    {Name: "ChevronLeft", SafeValue: 'ms-Icon--ChevronLeft'},
    {Name: "ChevronRight", SafeValue: 'ms-Icon--ChevronRight'},
    {Name: "Emoji2", SafeValue: 'ms-Icon--Emoji2'},
    {Name: "System", SafeValue: 'ms-Icon--System'},
    {Name: "Globe", SafeValue: 'ms-Icon--Globe'},
    {Name: "Unpin", SafeValue: 'ms-Icon--Unpin'},
    {Name: "Contact", SafeValue: 'ms-Icon--Contact'},
    {Name: "Memo", SafeValue: 'ms-Icon--Memo'},
    {Name: "WindowsLogo", SafeValue: 'ms-Icon--WindowsLogo'},
    {Name: "Error", SafeValue: 'ms-Icon--Error'},
    {Name: "Unlock", SafeValue: 'ms-Icon--Unlock'},
    {Name: "Calendar", SafeValue: 'ms-Icon--Calendar'},
    {Name: "Megaphone", SafeValue: 'ms-Icon--Megaphone'},
    {Name: "AutoEnhanceOn", SafeValue: 'ms-Icon--AutoEnhanceOn'},
    {Name: "AutoEnhanceOff", SafeValue: 'ms-Icon--AutoEnhanceOff'},
    {Name: "Color", SafeValue: 'ms-Icon--Color'},
    {Name: "SaveAs", SafeValue: 'ms-Icon--SaveAs'},
    {Name: "Light", SafeValue: 'ms-Icon--Light'},
    {Name: "Filters", SafeValue: 'ms-Icon--Filters'},
    {Name: "Contrast", SafeValue: 'ms-Icon--Contrast'},
    {Name: "Redo", SafeValue: 'ms-Icon--Redo'},
    {Name: "Undo", SafeValue: 'ms-Icon--Undo'},
    {Name: "Album", SafeValue: 'ms-Icon--Album'},
    {Name: "Rotate", SafeValue: 'ms-Icon--Rotate'},
    {Name: "PanoIndicator", SafeValue: 'ms-Icon--PanoIndicator'},
    {Name: "ThumbnailView", SafeValue: 'ms-Icon--ThumbnailView'},
    {Name: "Package", SafeValue: 'ms-Icon--Package'},
    {Name: "Warning", SafeValue: 'ms-Icon--Warning'},
    {Name: "Financial", SafeValue: 'ms-Icon--Financial'},
    {Name: "ShoppingCart", SafeValue: 'ms-Icon--ShoppingCart'},
    {Name: "Train", SafeValue: 'ms-Icon--Train'},
    {Name: "Flag", SafeValue: 'ms-Icon--Flag'},
    {Name: "Move", SafeValue: 'ms-Icon--Move'},
    {Name: "Page", SafeValue: 'ms-Icon--Page'},
    {Name: "TouchPointer", SafeValue: 'ms-Icon--TouchPointer'},
    {Name: "Merge", SafeValue: 'ms-Icon--Merge'},
    {Name: "TurnRight", SafeValue: 'ms-Icon--TurnRight'},
    {Name: "Ferry", SafeValue: 'ms-Icon--Ferry'},
    {Name: "Tab", SafeValue: 'ms-Icon--Tab'},
    {Name: "Admin", SafeValue: 'ms-Icon--Admin'},
    {Name: "TVMonitor", SafeValue: 'ms-Icon--TVMonitor'},
    {Name: "Speakers", SafeValue: 'ms-Icon--Speakers'},
    {Name: "Car", SafeValue: 'ms-Icon--Car'},
    {Name: "EatDrink", SafeValue: 'ms-Icon--EatDrink'},
    {Name: "LocationCircle", SafeValue: 'ms-Icon--LocationCircle'},
    {Name: "Home", SafeValue: 'ms-Icon--Home'},
    {Name: "SwitcherStartEnd", SafeValue: 'ms-Icon--SwitcherStartEnd'},
    {Name: "IncidentTriangle", SafeValue: 'ms-Icon--IncidentTriangle'},
    {Name: "Touch", SafeValue: 'ms-Icon--Touch'},
    {Name: "MapDirections", SafeValue: 'ms-Icon--MapDirections'},
    {Name: "History", SafeValue: 'ms-Icon--History'},
    {Name: "Location", SafeValue: 'ms-Icon--Location'},
    {Name: "Work", SafeValue: 'ms-Icon--Work'},
    {Name: "Recent", SafeValue: 'ms-Icon--Recent'},
    {Name: "Hotel", SafeValue: 'ms-Icon--Hotel'},
    {Name: "LocationDot", SafeValue: 'ms-Icon--LocationDot'},
    {Name: "News", SafeValue: 'ms-Icon--News'},
    {Name: "Chat", SafeValue: 'ms-Icon--Chat'},
    {Name: "Group", SafeValue: 'ms-Icon--Group'},
    {Name: "View", SafeValue: 'ms-Icon--View'},
    {Name: "Clear", SafeValue: 'ms-Icon--Clear'},
    {Name: "Sync", SafeValue: 'ms-Icon--Sync'},
    {Name: "Download", SafeValue: 'ms-Icon--Download'},
    {Name: "Help", SafeValue: 'ms-Icon--Help'},
    {Name: "Upload", SafeValue: 'ms-Icon--Upload'},
    {Name: "Emoji", SafeValue: 'ms-Icon--Emoji'},
    {Name: "MailForward", SafeValue: 'ms-Icon--MailForward'},
    {Name: "ClosePane", SafeValue: 'ms-Icon--ClosePane'},
    {Name: "OpenPane", SafeValue: 'ms-Icon--OpenPane'},
    {Name: "PreviewLink", SafeValue: 'ms-Icon--PreviewLink'},
    {Name: "ZoomIn", SafeValue: 'ms-Icon--ZoomIn'},
    {Name: "Bookmarks", SafeValue: 'ms-Icon--Bookmarks'},
    {Name: "Document", SafeValue: 'ms-Icon--Document'},
    {Name: "ProtectedDocument", SafeValue: 'ms-Icon--ProtectedDocument'},
    {Name: "OpenInNewWindow", SafeValue: 'ms-Icon--OpenInNewWindow'},
    {Name: "MailFill", SafeValue: 'ms-Icon--MailFill'},
    {Name: "ViewAll", SafeValue: 'ms-Icon--ViewAll'},
    {Name: "Switch", SafeValue: 'ms-Icon--Switch'},
    {Name: "Rename", SafeValue: 'ms-Icon--Rename'},
    {Name: "Folder", SafeValue: 'ms-Icon--Folder'},
    {Name: "Picture", SafeValue: 'ms-Icon--Picture'},
    {Name: "ShowResults", SafeValue: 'ms-Icon--ShowResults'},
    {Name: "Message", SafeValue: 'ms-Icon--Message'},
    {Name: "CalendarDay", SafeValue: 'ms-Icon--CalendarDay'},
    {Name: "CalendarWeek", SafeValue: 'ms-Icon--CalendarWeek'},
    {Name: "MailReplyAll", SafeValue: 'ms-Icon--MailReplyAll'},
    {Name: "Read", SafeValue: 'ms-Icon--Read'},
    {Name: "PaymentCard", SafeValue: 'ms-Icon--PaymentCard'},
    {Name: "Copy", SafeValue: 'ms-Icon--Copy'},
    {Name: "Important", SafeValue: 'ms-Icon--Important'},
    {Name: "MailReply", SafeValue: 'ms-Icon--MailReply'},
    {Name: "Sort", SafeValue: 'ms-Icon--Sort'},
    {Name: "GotoToday", SafeValue: 'ms-Icon--GotoToday'},
    {Name: "Font", SafeValue: 'ms-Icon--Font'},
    {Name: "FontColor", SafeValue: 'ms-Icon--FontColor'},
    {Name: "FolderFill", SafeValue: 'ms-Icon--FolderFill'},
    {Name: "Permissions", SafeValue: 'ms-Icon--Permissions'},
    {Name: "DisableUpdates", SafeValue: 'ms-Icon--DisableUpdates'},
    {Name: "Unfavorite", SafeValue: 'ms-Icon--Unfavorite'},
    {Name: "Italic", SafeValue: 'ms-Icon--Italic'},
    {Name: "Underline", SafeValue: 'ms-Icon--Underline'},
    {Name: "Bold", SafeValue: 'ms-Icon--Bold'},
    {Name: "MoveToFolder", SafeValue: 'ms-Icon--MoveToFolder'},
    {Name: "Dislike", SafeValue: 'ms-Icon--Dislike'},
    {Name: "Like", SafeValue: 'ms-Icon--Like'},
    {Name: "AlignRight", SafeValue: 'ms-Icon--AlignRight'},
    {Name: "AlignCenter", SafeValue: 'ms-Icon--AlignCenter'},
    {Name: "AlignLeft", SafeValue: 'ms-Icon--AlignLeft'},
    {Name: "OpenFile", SafeValue: 'ms-Icon--OpenFile'},
    {Name: "FontDecrease", SafeValue: 'ms-Icon--FontDecrease'},
    {Name: "FontIncrease", SafeValue: 'ms-Icon--FontIncrease'},
    {Name: "FontSize", SafeValue: 'ms-Icon--FontSize'},
    {Name: "CellPhone", SafeValue: 'ms-Icon--CellPhone'},
    {Name: "Tag", SafeValue: 'ms-Icon--Tag'},
    {Name: "Library", SafeValue: 'ms-Icon--Library'},
    {Name: "PostUpdate", SafeValue: 'ms-Icon--PostUpdate'},
    {Name: "NewFolder", SafeValue: 'ms-Icon--NewFolder'},
    {Name: "CalendarReply", SafeValue: 'ms-Icon--CalendarReply'},
    {Name: "UnsyncFolder", SafeValue: 'ms-Icon--UnsyncFolder'},
    {Name: "SyncFolder", SafeValue: 'ms-Icon--SyncFolder'},
    {Name: "BlockContact", SafeValue: 'ms-Icon--BlockContact'},
    {Name: "AddFriend", SafeValue: 'ms-Icon--AddFriend'},
    {Name: "BulletedList", SafeValue: 'ms-Icon--BulletedList'},
    {Name: "Preview", SafeValue: 'ms-Icon--Preview'},
    {Name: "DockLeft", SafeValue: 'ms-Icon--DockLeft'},
    {Name: "DockRight", SafeValue: 'ms-Icon--DockRight'},
    {Name: "Repair", SafeValue: 'ms-Icon--Repair'},
    {Name: "Accounts", SafeValue: 'ms-Icon--Accounts'},
    {Name: "RadioBullet", SafeValue: 'ms-Icon--RadioBullet'},
    {Name: "Stopwatch", SafeValue: 'ms-Icon--Stopwatch'},
    {Name: "Clock", SafeValue: 'ms-Icon--Clock'},
    {Name: "AlarmClock", SafeValue: 'ms-Icon--AlarmClock'},
    {Name: "Hospital", SafeValue: 'ms-Icon--Hospital'},
    {Name: "Timer", SafeValue: 'ms-Icon--Timer'},
    {Name: "FullCircleMask", SafeValue: 'ms-Icon--FullCircleMask'},
    {Name: "LocationFill", SafeValue: 'ms-Icon--LocationFill'},
    {Name: "ChromeMinimize", SafeValue: 'ms-Icon--ChromeMinimize'},
    {Name: "Annotation", SafeValue: 'ms-Icon--Annotation'},
    {Name: "ChromeClose", SafeValue: 'ms-Icon--ChromeClose'},
    {Name: "Accept", SafeValue: 'ms-Icon--Accept'},
    {Name: "Fingerprint", SafeValue: 'ms-Icon--Fingerprint'},
    {Name: "Handwriting", SafeValue: 'ms-Icon--Handwriting'},
    {Name: "StackIndicator", SafeValue: 'ms-Icon--StackIndicator'},
    {Name: "Completed", SafeValue: 'ms-Icon--Completed'},
    {Name: "Label", SafeValue: 'ms-Icon--Label'},
    {Name: "FlickDown", SafeValue: 'ms-Icon--FlickDown'},
    {Name: "FlickUp", SafeValue: 'ms-Icon--FlickUp'},
    {Name: "FlickLeft", SafeValue: 'ms-Icon--FlickLeft'},
    {Name: "FlickRight", SafeValue: 'ms-Icon--FlickRight'},
    {Name: "MusicInCollection", SafeValue: 'ms-Icon--MusicInCollection'},
    {Name: "OneDrive", SafeValue: 'ms-Icon--OneDrive'},
    {Name: "CompassNW", SafeValue: 'ms-Icon--CompassNW'},
    {Name: "Code", SafeValue: 'ms-Icon--Code'},
    {Name: "LightningBolt", SafeValue: 'ms-Icon--LightningBolt'},
    {Name: "Info", SafeValue: 'ms-Icon--Info'},
    {Name: "CalculatorAddition", SafeValue: 'ms-Icon--CalculatorAddition'},
    {Name: "CalculatorSubtract", SafeValue: 'ms-Icon--CalculatorSubtract'},
    {Name: "PrintfaxPrinterFile", SafeValue: 'ms-Icon--PrintfaxPrinterFile'},
    {Name: "Health", SafeValue: 'ms-Icon--Health'},
    {Name: "ChevronUpSmall", SafeValue: 'ms-Icon--ChevronUpSmall'},
    {Name: "ChevronDownSmall", SafeValue: 'ms-Icon--ChevronDownSmall'},
    {Name: "ChevronLeftSmall", SafeValue: 'ms-Icon--ChevronLeftSmall'},
    {Name: "ChevronRightSmall", SafeValue: 'ms-Icon--ChevronRightSmall'},
    {Name: "ChevronUpMed", SafeValue: 'ms-Icon--ChevronUpMed'},
    {Name: "ChevronDownMed", SafeValue: 'ms-Icon--ChevronDownMed'},
    {Name: "ChevronLeftMed", SafeValue: 'ms-Icon--ChevronLeftMed'},
    {Name: "ChevronRightMed", SafeValue: 'ms-Icon--ChevronRightMed'},
    {Name: "Dictionary", SafeValue: 'ms-Icon--Dictionary'},
    {Name: "ChromeBack", SafeValue: 'ms-Icon--ChromeBack'},
    {Name: "PC1", SafeValue: 'ms-Icon--PC1'},
    {Name: "PresenceChickletVideo", SafeValue: 'ms-Icon--PresenceChickletVideo'},
    {Name: "Reply", SafeValue: 'ms-Icon--Reply'},
    {Name: "DoubleChevronLeftMed", SafeValue: 'ms-Icon--DoubleChevronLeftMed'},
    {Name: "Volume0", SafeValue: 'ms-Icon--Volume0'},
    {Name: "Volume1", SafeValue: 'ms-Icon--Volume1'},
    {Name: "Volume2", SafeValue: 'ms-Icon--Volume2'},
    {Name: "Volume3", SafeValue: 'ms-Icon--Volume3'},
    {Name: "CaretHollow", SafeValue: 'ms-Icon--CaretHollow'},
    {Name: "CaretSolid", SafeValue: 'ms-Icon--CaretSolid'},
    {Name: "Pinned", SafeValue: 'ms-Icon--Pinned'},
    {Name: "PinnedFill", SafeValue: 'ms-Icon--PinnedFill'},
    {Name: "Chart", SafeValue: 'ms-Icon--Chart'},
    {Name: "BidiLtr", SafeValue: 'ms-Icon--BidiLtr'},
    {Name: "BidiRtl", SafeValue: 'ms-Icon--BidiRtl'},
    {Name: "RevToggleKey", SafeValue: 'ms-Icon--RevToggleKey'},
    {Name: "RightDoubleQuote", SafeValue: 'ms-Icon--RightDoubleQuote'},
    {Name: "Sunny", SafeValue: 'ms-Icon--Sunny'},
    {Name: "CloudWeather", SafeValue: 'ms-Icon--CloudWeather'},
    {Name: "Cloudy", SafeValue: 'ms-Icon--Cloudy'},
    {Name: "PartlyCloudyDay", SafeValue: 'ms-Icon--PartlyCloudyDay'},
    {Name: "PartlyCloudyNight", SafeValue: 'ms-Icon--PartlyCloudyNight'},
    {Name: "ClearNight", SafeValue: 'ms-Icon--ClearNight'},
    {Name: "RainShowersDay", SafeValue: 'ms-Icon--RainShowersDay'},
    {Name: "Rain", SafeValue: 'ms-Icon--Rain'},
    {Name: "RainSnow", SafeValue: 'ms-Icon--RainSnow'},
    {Name: "Snow", SafeValue: 'ms-Icon--Snow'},
    {Name: "BlowingSnow", SafeValue: 'ms-Icon--BlowingSnow'},
    {Name: "Frigid", SafeValue: 'ms-Icon--Frigid'},
    {Name: "Fog", SafeValue: 'ms-Icon--Fog'},
    {Name: "Squalls", SafeValue: 'ms-Icon--Squalls'},
    {Name: "Duststorm", SafeValue: 'ms-Icon--Duststorm'},
    {Name: "Precipitation", SafeValue: 'ms-Icon--Precipitation'},
    {Name: "Ringer", SafeValue: 'ms-Icon--Ringer'},
    {Name: "PDF", SafeValue: 'ms-Icon--PDF'},
    {Name: "SortLines", SafeValue: 'ms-Icon--SortLines'},
    {Name: "Ribbon", SafeValue: 'ms-Icon--Ribbon'},
    {Name: "CheckList", SafeValue: 'ms-Icon--CheckList'},
    {Name: "Generate", SafeValue: 'ms-Icon--Generate'},
    {Name: "Equalizer", SafeValue: 'ms-Icon--Equalizer'},
    {Name: "BarChartHorizontal", SafeValue: 'ms-Icon--BarChartHorizontal'},
    {Name: "Freezing", SafeValue: 'ms-Icon--Freezing'},
    {Name: "SnowShowerDay", SafeValue: 'ms-Icon--SnowShowerDay'},
    {Name: "HailDay", SafeValue: 'ms-Icon--HailDay'},
    {Name: "WorkFlow", SafeValue: 'ms-Icon--WorkFlow'},
    {Name: "StoreLogoMed", SafeValue: 'ms-Icon--StoreLogoMed'},
    {Name: "RainShowersNight", SafeValue: 'ms-Icon--RainShowersNight'},
    {Name: "SnowShowerNight", SafeValue: 'ms-Icon--SnowShowerNight'},
    {Name: "HailNight", SafeValue: 'ms-Icon--HailNight'},
    {Name: "Info2", SafeValue: 'ms-Icon--Info2'},
    {Name: "StoreLogo", SafeValue: 'ms-Icon--StoreLogo'},
    {Name: "Broom", SafeValue: 'ms-Icon--Broom'},
    {Name: "MusicInCollectionFill", SafeValue: 'ms-Icon--MusicInCollectionFill'},
    {Name: "List", SafeValue: 'ms-Icon--List'},
    {Name: "Asterisk", SafeValue: 'ms-Icon--Asterisk'},
    {Name: "ErrorBadge", SafeValue: 'ms-Icon--ErrorBadge'},
    {Name: "CircleRing", SafeValue: 'ms-Icon--CircleRing'},
    {Name: "CircleFill", SafeValue: 'ms-Icon--CircleFill'},
    {Name: "Lightbulb", SafeValue: 'ms-Icon--Lightbulb'},
    {Name: "StatusTriangle", SafeValue: 'ms-Icon--StatusTriangle'},
    {Name: "VolumeDisabled", SafeValue: 'ms-Icon--VolumeDisabled'},
    {Name: "Puzzle", SafeValue: 'ms-Icon--Puzzle'},
    {Name: "EmojiNeutral", SafeValue: 'ms-Icon--EmojiNeutral'},
    {Name: "EmojiDisappointed", SafeValue: 'ms-Icon--EmojiDisappointed'},
    {Name: "HomeSolid", SafeValue: 'ms-Icon--HomeSolid'},
    {Name: "Cocktails", SafeValue: 'ms-Icon--Cocktails'},
    {Name: "Articles", SafeValue: 'ms-Icon--Articles'},
    {Name: "Cycling", SafeValue: 'ms-Icon--Cycling'},
    {Name: "DietPlanNotebook", SafeValue: 'ms-Icon--DietPlanNotebook'},
    {Name: "Pill", SafeValue: 'ms-Icon--Pill'},
    {Name: "Running", SafeValue: 'ms-Icon--Running'},
    {Name: "Weights", SafeValue: 'ms-Icon--Weights'},
    {Name: "BarChart4", SafeValue: 'ms-Icon--BarChart4'},
    {Name: "CirclePlus", SafeValue: 'ms-Icon--CirclePlus'},
    {Name: "Coffee", SafeValue: 'ms-Icon--Coffee'},
    {Name: "Cotton", SafeValue: 'ms-Icon--Cotton'},
    {Name: "Market", SafeValue: 'ms-Icon--Market'},
    {Name: "Money", SafeValue: 'ms-Icon--Money'},
    {Name: "PieDouble", SafeValue: 'ms-Icon--PieDouble'},
    {Name: "RemoveFilter", SafeValue: 'ms-Icon--RemoveFilter'},
    {Name: "StockDown", SafeValue: 'ms-Icon--StockDown'},
    {Name: "StockUp", SafeValue: 'ms-Icon--StockUp'},
    {Name: "Cricket", SafeValue: 'ms-Icon--Cricket'},
    {Name: "Golf", SafeValue: 'ms-Icon--Golf'},
    {Name: "Baseball", SafeValue: 'ms-Icon--Baseball'},
    {Name: "Soccer", SafeValue: 'ms-Icon--Soccer'},
    {Name: "MoreSports", SafeValue: 'ms-Icon--MoreSports'},
    {Name: "AutoRacing", SafeValue: 'ms-Icon--AutoRacing'},
    {Name: "CollegeHoops", SafeValue: 'ms-Icon--CollegeHoops'},
    {Name: "CollegeFootball", SafeValue: 'ms-Icon--CollegeFootball'},
    {Name: "ProFootball", SafeValue: 'ms-Icon--ProFootball'},
    {Name: "ProHockey", SafeValue: 'ms-Icon--ProHockey'},
    {Name: "Rugby", SafeValue: 'ms-Icon--Rugby'},
    {Name: "Tennis", SafeValue: 'ms-Icon--Tennis'},
    {Name: "Arrivals", SafeValue: 'ms-Icon--Arrivals'},
    {Name: "Design", SafeValue: 'ms-Icon--Design'},
    {Name: "Website", SafeValue: 'ms-Icon--Website'},
    {Name: "Drop", SafeValue: 'ms-Icon--Drop'},
    {Name: "Snow", SafeValue: 'ms-Icon--Snow'},
    {Name: "BusSolid", SafeValue: 'ms-Icon--BusSolid'},
    {Name: "FerrySolid", SafeValue: 'ms-Icon--FerrySolid'},
    {Name: "TrainSolid", SafeValue: 'ms-Icon--TrainSolid'},
    {Name: "Heart", SafeValue: 'ms-Icon--Heart'},
    {Name: "HeartFill", SafeValue: 'ms-Icon--HeartFill'},
    {Name: "Ticket", SafeValue: 'ms-Icon--Ticket'},
    {Name: "AzureLogo", SafeValue: 'ms-Icon--AzureLogo'},
    {Name: "BingLogo", SafeValue: 'ms-Icon--BingLogo'},
    {Name: "MSNLogo", SafeValue: 'ms-Icon--MSNLogo'},
    {Name: "OutlookLogo", SafeValue: 'ms-Icon--OutlookLogo'},
    {Name: "OfficeLogo", SafeValue: 'ms-Icon--OfficeLogo'},
    {Name: "SkypeLogo", SafeValue: 'ms-Icon--SkypeLogo'},
    {Name: "Door", SafeValue: 'ms-Icon--Door'},
    {Name: "GiftCard", SafeValue: 'ms-Icon--GiftCard'},
    {Name: "DoubleBookmark", SafeValue: 'ms-Icon--DoubleBookmark'},
    {Name: "StatusErrorFull", SafeValue: 'ms-Icon--StatusErrorFull'},
    {Name: "Certificate", SafeValue: 'ms-Icon--Certificate'},
    {Name: "Photo2", SafeValue: 'ms-Icon--Photo2'},
    {Name: "CloudDownload", SafeValue: 'ms-Icon--CloudDownload'},
    {Name: "WindDirection", SafeValue: 'ms-Icon--WindDirection'},
    {Name: "Family", SafeValue: 'ms-Icon--Family'},
    {Name: "CSS", SafeValue: 'ms-Icon--CSS'},
    {Name: "JS", SafeValue: 'ms-Icon--JS'},
    {Name: "ReminderGroup", SafeValue: 'ms-Icon--ReminderGroup'},
    {Name: "Section", SafeValue: 'ms-Icon--Section'},
    {Name: "OneNoteLogo", SafeValue: 'ms-Icon--OneNoteLogo'},
    {Name: "ToggleFilled", SafeValue: 'ms-Icon--ToggleFilled'},
    {Name: "ToggleBorder", SafeValue: 'ms-Icon--ToggleBorder'},
    {Name: "SliderThumb", SafeValue: 'ms-Icon--SliderThumb'},
    {Name: "ToggleThumb", SafeValue: 'ms-Icon--ToggleThumb'},
    {Name: "Documentation", SafeValue: 'ms-Icon--Documentation'},
    {Name: "Badge", SafeValue: 'ms-Icon--Badge'},
    {Name: "Giftbox", SafeValue: 'ms-Icon--Giftbox'},
    {Name: "ExcelLogo", SafeValue: 'ms-Icon--ExcelLogo'},
    {Name: "WordLogo", SafeValue: 'ms-Icon--WordLogo'},
    {Name: "PowerPointLogo", SafeValue: 'ms-Icon--PowerPointLogo'},
    {Name: "Cafe", SafeValue: 'ms-Icon--Cafe'},
    {Name: "SpeedHigh", SafeValue: 'ms-Icon--SpeedHigh'},
    {Name: "MusicNote", SafeValue: 'ms-Icon--MusicNote'},
    {Name: "EdgeLogo", SafeValue: 'ms-Icon--EdgeLogo'},
    {Name: "CompletedSolid", SafeValue: 'ms-Icon--CompletedSolid'},
    {Name: "AlbumRemove", SafeValue: 'ms-Icon--AlbumRemove'},
    {Name: "MessageFill", SafeValue: 'ms-Icon--MessageFill'},
    {Name: "TabletSelected", SafeValue: 'ms-Icon--TabletSelected'},
    {Name: "MobileSelected", SafeValue: 'ms-Icon--MobileSelected'},
    {Name: "LaptopSelected", SafeValue: 'ms-Icon--LaptopSelected'},
    {Name: "TVMonitorSelected", SafeValue: 'ms-Icon--TVMonitorSelected'},
    {Name: "DeveloperTools", SafeValue: 'ms-Icon--DeveloperTools'},
    {Name: "InsertTextBox", SafeValue: 'ms-Icon--InsertTextBox'},
    {Name: "LowerBrightness", SafeValue: 'ms-Icon--LowerBrightness'},
    {Name: "CloudUpload", SafeValue: 'ms-Icon--CloudUpload'},
    {Name: "DateTime", SafeValue: 'ms-Icon--DateTime'},
    {Name: "Tiles", SafeValue: 'ms-Icon--Tiles'},
    {Name: "Org", SafeValue: 'ms-Icon--Org'},
    {Name: "PartyLeader", SafeValue: 'ms-Icon--PartyLeader'},
    {Name: "DRM", SafeValue: 'ms-Icon--DRM'},
    {Name: "CloudAdd", SafeValue: 'ms-Icon--CloudAdd'},
    {Name: "AppIconDefault", SafeValue: 'ms-Icon--AppIconDefault'},
    {Name: "Photo2Add", SafeValue: 'ms-Icon--Photo2Add'},
    {Name: "Photo2Remove", SafeValue: 'ms-Icon--Photo2Remove'},
    {Name: "POI", SafeValue: 'ms-Icon--POI'},
    {Name: "FacebookLogo", SafeValue: 'ms-Icon--FacebookLogo'},
    {Name: "AddTo", SafeValue: 'ms-Icon--AddTo'},
    {Name: "RadioBtnOn", SafeValue: 'ms-Icon--RadioBtnOn'},
    {Name: "Embed", SafeValue: 'ms-Icon--Embed'},
    {Name: "VideoSolid", SafeValue: 'ms-Icon--VideoSolid'},
    {Name: "Teamwork", SafeValue: 'ms-Icon--Teamwork'},
    {Name: "PeopleAdd", SafeValue: 'ms-Icon--PeopleAdd'},
    {Name: "Glasses", SafeValue: 'ms-Icon--Glasses'},
    {Name: "DateTime2", SafeValue: 'ms-Icon--DateTime2'},
    {Name: "Shield", SafeValue: 'ms-Icon--Shield'},
    {Name: "Header1", SafeValue: 'ms-Icon--Header1'},
    {Name: "PageAdd", SafeValue: 'ms-Icon--PageAdd'},
    {Name: "NumberedList", SafeValue: 'ms-Icon--NumberedList'},
    {Name: "PowerBILogo", SafeValue: 'ms-Icon--PowerBILogo'},
    {Name: "Product", SafeValue: 'ms-Icon--Product'},
    {Name: "Blocked2", SafeValue: 'ms-Icon--Blocked2'},
    {Name: "FangBody", SafeValue: 'ms-Icon--FangBody'},
    {Name: "Glimmer", SafeValue: 'ms-Icon--Glimmer'},
    {Name: "ChatInviteFriend", SafeValue: 'ms-Icon--ChatInviteFriend'},
    {Name: "SharepointLogo", SafeValue: 'ms-Icon--SharepointLogo'},
    {Name: "YammerLogo", SafeValue: 'ms-Icon--YammerLogo'},
    {Name: "ReturnToSession", SafeValue: 'ms-Icon--ReturnToSession'},
    {Name: "OpenFolderHorizontal", SafeValue: 'ms-Icon--OpenFolderHorizontal'},
    {Name: "SwayLogo", SafeValue: 'ms-Icon--SwayLogo'},
    {Name: "OutOfOffice", SafeValue: 'ms-Icon--OutOfOffice'},
    {Name: "Trophy", SafeValue: 'ms-Icon--Trophy'},
    {Name: "ReopenPages", SafeValue: 'ms-Icon--ReopenPages'},
    {Name: "AADLogo", SafeValue: 'ms-Icon--AADLogo'},
    {Name: "AccessLogo", SafeValue: 'ms-Icon--AccessLogo'},
    {Name: "AdminALogo", SafeValue: 'ms-Icon--AdminALogo'},
    {Name: "AdminCLogo", SafeValue: 'ms-Icon--AdminCLogo'},
    {Name: "AdminDLogo", SafeValue: 'ms-Icon--AdminDLogo'},
    {Name: "AdminELogo", SafeValue: 'ms-Icon--AdminELogo'},
    {Name: "AdminLLogo", SafeValue: 'ms-Icon--AdminLLogo'},
    {Name: "AdminMLogo", SafeValue: 'ms-Icon--AdminMLogo'},
    {Name: "AdminOLogo", SafeValue: 'ms-Icon--AdminOLogo'},
    {Name: "AdminPLogo", SafeValue: 'ms-Icon--AdminPLogo'},
    {Name: "AdminSLogo", SafeValue: 'ms-Icon--AdminSLogo'},
    {Name: "AdminYLogo", SafeValue: 'ms-Icon--AdminYLogo'},
    {Name: "AlchemyLogo", SafeValue: 'ms-Icon--AlchemyLogo'},
    {Name: "BoxLogo", SafeValue: 'ms-Icon--BoxLogo'},
    {Name: "DelveLogo", SafeValue: 'ms-Icon--DelveLogo'},
    {Name: "DropboxLogo", SafeValue: 'ms-Icon--DropboxLogo'},
    {Name: "ExchangeLogo", SafeValue: 'ms-Icon--ExchangeLogo'},
    {Name: "LyncLogo", SafeValue: 'ms-Icon--LyncLogo'},
    {Name: "OfficeVideoLogo", SafeValue: 'ms-Icon--OfficeVideoLogo'},
    {Name: "ParatureLogo", SafeValue: 'ms-Icon--ParatureLogo'},
    {Name: "SocialListeningLogo", SafeValue: 'ms-Icon--SocialListeningLogo'},
    {Name: "VisioLogo", SafeValue: 'ms-Icon--VisioLogo'},
    {Name: "Balloons", SafeValue: 'ms-Icon--Balloons'},
    {Name: "Cat", SafeValue: 'ms-Icon--Cat'},
    {Name: "MailAlert", SafeValue: 'ms-Icon--MailAlert'},
    {Name: "MailCheck", SafeValue: 'ms-Icon--MailCheck'},
    {Name: "MailLowImportance", SafeValue: 'ms-Icon--MailLowImportance'},
    {Name: "MailPause", SafeValue: 'ms-Icon--MailPause'},
    {Name: "MailRepeat", SafeValue: 'ms-Icon--MailRepeat'},
    {Name: "SecurityGroup", SafeValue: 'ms-Icon--SecurityGroup'},
    {Name: "VoicemailForward", SafeValue: 'ms-Icon--VoicemailForward'},
    {Name: "VoicemailReply", SafeValue: 'ms-Icon--VoicemailReply'},
    {Name: "Waffle", SafeValue: 'ms-Icon--Waffle'},
    {Name: "RemoveEvent", SafeValue: 'ms-Icon--RemoveEvent'},
    {Name: "EventInfo", SafeValue: 'ms-Icon--EventInfo'},
    {Name: "ForwardEvent", SafeValue: 'ms-Icon--ForwardEvent'},
    {Name: "WipePhone", SafeValue: 'ms-Icon--WipePhone'},
    {Name: "AddOnlineMeeting", SafeValue: 'ms-Icon--AddOnlineMeeting'},
    {Name: "JoinOnlineMeeting", SafeValue: 'ms-Icon--JoinOnlineMeeting'},
    {Name: "RemoveLink", SafeValue: 'ms-Icon--RemoveLink'},
    {Name: "PeopleBlock", SafeValue: 'ms-Icon--PeopleBlock'},
    {Name: "PeopleRepeat", SafeValue: 'ms-Icon--PeopleRepeat'},
    {Name: "PeopleAlert", SafeValue: 'ms-Icon--PeopleAlert'},
    {Name: "PeoplePause", SafeValue: 'ms-Icon--PeoplePause'},
    {Name: "TransferCall", SafeValue: 'ms-Icon--TransferCall'},
    {Name: "AddPhone", SafeValue: 'ms-Icon--AddPhone'},
    {Name: "UnknownCall", SafeValue: 'ms-Icon--UnknownCall'},
    {Name: "NoteReply", SafeValue: 'ms-Icon--NoteReply'},
    {Name: "NoteForward", SafeValue: 'ms-Icon--NoteForward'},
    {Name: "NotePinned", SafeValue: 'ms-Icon--NotePinned'},
    {Name: "RemoveOccurrence", SafeValue: 'ms-Icon--RemoveOccurrence'},
    {Name: "Timeline", SafeValue: 'ms-Icon--Timeline'},
    {Name: "EditNote", SafeValue: 'ms-Icon--EditNote'},
    {Name: "CircleHalfFull", SafeValue: 'ms-Icon--CircleHalfFull'},
    {Name: "Room", SafeValue: 'ms-Icon--Room'},
    {Name: "Unsubscribe", SafeValue: 'ms-Icon--Unsubscribe'},
    {Name: "Subscribe", SafeValue: 'ms-Icon--Subscribe'},
    {Name: "RecurringTask", SafeValue: 'ms-Icon--RecurringTask'},
    {Name: "TaskManager", SafeValue: 'ms-Icon--TaskManager'},
    {Name: "Combine", SafeValue: 'ms-Icon--Combine'},
    {Name: "Split", SafeValue: 'ms-Icon--Split'},
    {Name: "DoubleChevronUp", SafeValue: 'ms-Icon--DoubleChevronUp'},
    {Name: "DoubleChevronLeft", SafeValue: 'ms-Icon--DoubleChevronLeft'},
    {Name: "DoubleChevronRight", SafeValue: 'ms-Icon--DoubleChevronRight'},
    {Name: "Ascending", SafeValue: 'ms-Icon--Ascending'},
    {Name: "Descending", SafeValue: 'ms-Icon--Descending'},
    {Name: "TextBox", SafeValue: 'ms-Icon--TextBox'},
    {Name: "TextField", SafeValue: 'ms-Icon--TextField'},
    {Name: "NumberField", SafeValue: 'ms-Icon--NumberField'},
    {Name: "Dropdown", SafeValue: 'ms-Icon--Dropdown'},
    {Name: "BookingsLogo", SafeValue: 'ms-Icon--BookingsLogo'},
    {Name: "ClassNotebookLogo", SafeValue: 'ms-Icon--ClassNotebookLogo'},
    {Name: "CollabsDBLogo", SafeValue: 'ms-Icon--CollabsDBLogo'},
    {Name: "DelveAnalyticsLogo", SafeValue: 'ms-Icon--DelveAnalyticsLogo'},
    {Name: "DocsLogo", SafeValue: 'ms-Icon--DocsLogo'},
    {Name: "DynamicsCRMLogo", SafeValue: 'ms-Icon--DynamicsCRMLogo'},
    {Name: "DynamicSMBLogo", SafeValue: 'ms-Icon--DynamicSMBLogo'},
    {Name: "OfficeAssistantLogo", SafeValue: 'ms-Icon--OfficeAssistantLogo'},
    {Name: "OfficeStoreLogo", SafeValue: 'ms-Icon--OfficeStoreLogo'},
    {Name: "OneNoteEduLogo", SafeValue: 'ms-Icon--OneNoteEduLogo'},
    {Name: "Planner", SafeValue: 'ms-Icon--Planner'},
    {Name: "PowerApps", SafeValue: 'ms-Icon--PowerApps'},
    {Name: "Suitcase", SafeValue: 'ms-Icon--Suitcase'},
    {Name: "ProjectLogo", SafeValue: 'ms-Icon--ProjectLogo'},
    {Name: "CaretLeft8", SafeValue: 'ms-Icon--CaretLeft8'},
    {Name: "CaretRight8", SafeValue: 'ms-Icon--CaretRight8'},
    {Name: "CaretUp8", SafeValue: 'ms-Icon--CaretUp8'},
    {Name: "CaretDown8", SafeValue: 'ms-Icon--CaretDown8'},
    {Name: "CaretLeftSolid8", SafeValue: 'ms-Icon--CaretLeftSolid8'},
    {Name: "CaretRightSolid8", SafeValue: 'ms-Icon--CaretRightSolid8'},
    {Name: "CaretUpSolid8", SafeValue: 'ms-Icon--CaretUpSolid8'},
    {Name: "CaretDownSolid8", SafeValue: 'ms-Icon--CaretDownSolid8'},
    {Name: "ClearFormatting", SafeValue: 'ms-Icon--ClearFormatting'},
    {Name: "Superscript", SafeValue: 'ms-Icon--Superscript'},
    {Name: "Subscript", SafeValue: 'ms-Icon--Subscript'},
    {Name: "Strikethrough", SafeValue: 'ms-Icon--Strikethrough'},
    {Name: "SingleBookmark", SafeValue: 'ms-Icon--SingleBookmark'},
    {Name: "DoubleChevronDown", SafeValue: 'ms-Icon--DoubleChevronDown'},
    {Name: "ReplyAll", SafeValue: 'ms-Icon--ReplyAll'},
    {Name: "GoogleDriveLogo", SafeValue: 'ms-Icon--GoogleDriveLogo'},
    {Name: "Questionnaire", SafeValue: 'ms-Icon--Questionnaire'},
    {Name: "AddGroup", SafeValue: 'ms-Icon--AddGroup'},
    {Name: "TemporaryUser", SafeValue: 'ms-Icon--TemporaryUser'},
    {Name: "GroupedDescending", SafeValue: 'ms-Icon--GroupedDescending'},
    {Name: "GroupedAscending", SafeValue: 'ms-Icon--GroupedAscending'},
    {Name: "SortUp", SafeValue: 'ms-Icon--SortUp'},
    {Name: "SortDown", SafeValue: 'ms-Icon--SortDown'},
    {Name: "AwayStatus", SafeValue: 'ms-Icon--AwayStatus'},
    {Name: "SyncToPC", SafeValue: 'ms-Icon--SyncToPC'},
    {Name: "AustralianRules", SafeValue: 'ms-Icon--AustralianRules'},
    {Name: "DoubleChevronUp12", SafeValue: 'ms-Icon--DoubleChevronUp12'},
    {Name: "DoubleChevronDown12", SafeValue: 'ms-Icon--DoubleChevronDown12'},
    {Name: "DoubleChevronLeft12", SafeValue: 'ms-Icon--DoubleChevronLeft12'},
    {Name: "DoubleChevronRight12", SafeValue: 'ms-Icon--DoubleChevronRight12'},
    {Name: "CalendarAgenda", SafeValue: 'ms-Icon--CalendarAgenda'},
    {Name: "AddEvent", SafeValue: 'ms-Icon--AddEvent'},
    {Name: "AssetLibrary", SafeValue: 'ms-Icon--AssetLibrary'},
    {Name: "DataConnectionLibrary", SafeValue: 'ms-Icon--DataConnectionLibrary'},
    {Name: "DocLibrary", SafeValue: 'ms-Icon--DocLibrary'},
    {Name: "FormLibrary", SafeValue: 'ms-Icon--FormLibrary'},
    {Name: "ReportLibrary", SafeValue: 'ms-Icon--ReportLibrary'},
    {Name: "ContactCard", SafeValue: 'ms-Icon--ContactCard'},
    {Name: "CustomList", SafeValue: 'ms-Icon--CustomList'},
    {Name: "IssueTracking", SafeValue: 'ms-Icon--IssueTracking'},
    {Name: "PictureLibrary", SafeValue: 'ms-Icon--PictureLibrary'},
    {Name: "AppForOfficeLogo", SafeValue: 'ms-Icon--AppForOfficeLogo'},
    {Name: "OfflineOneDriveParachute", SafeValue: 'ms-Icon--OfflineOneDriveParachute'},
    {Name: "OfflineOneDriveParachuteDisabled", SafeValue: 'ms-Icon--OfflineOneDriveParachuteDisabled'},
    {Name: "LargeGrid", SafeValue: 'ms-Icon--LargeGrid'},
    {Name: "TriangleSolidUp12", SafeValue: 'ms-Icon--TriangleSolidUp12'},
    {Name: "TriangleSolidDown12", SafeValue: 'ms-Icon--TriangleSolidDown12'},
    {Name: "TriangleSolidLeft12", SafeValue: 'ms-Icon--TriangleSolidLeft12'},
    {Name: "TriangleSolidRight12", SafeValue: 'ms-Icon--TriangleSolidRight12'},
    {Name: "TriangleUp12", SafeValue: 'ms-Icon--TriangleUp12'},
    {Name: "TriangleDown12", SafeValue: 'ms-Icon--TriangleDown12'},
    {Name: "TriangleLeft12", SafeValue: 'ms-Icon--TriangleLeft12'},
    {Name: "TriangleRight12", SafeValue: 'ms-Icon--TriangleRight12'},
    {Name: "ArrowUpRight8", SafeValue: 'ms-Icon--ArrowUpRight8'},
    {Name: "ArrowDownRight8", SafeValue: 'ms-Icon--ArrowDownRight8'},
    {Name: "DocumentSet", SafeValue: 'ms-Icon--DocumentSet'},
    {Name: "DelveAnalytics", SafeValue: 'ms-Icon--DelveAnalytics'},
    {Name: "OneDriveAdd", SafeValue: 'ms-Icon--OneDriveAdd'},
    {Name: "Header2", SafeValue: 'ms-Icon--Header2'},
    {Name: "Header3", SafeValue: 'ms-Icon--Header3'},
    {Name: "Header4", SafeValue: 'ms-Icon--Header4'},
    {Name: "MarketDown", SafeValue: 'ms-Icon--MarketDown'},
    {Name: "CalendarWorkWeek", SafeValue: 'ms-Icon--CalendarWorkWeek'},
    {Name: "SidePanel", SafeValue: 'ms-Icon--SidePanel'},
    {Name: "GlobeFavorite", SafeValue: 'ms-Icon--GlobeFavorite'},
    {Name: "CaretTopLeftSolid8", SafeValue: 'ms-Icon--CaretTopLeftSolid8'},
    {Name: "CaretTopRightSolid8", SafeValue: 'ms-Icon--CaretTopRightSolid8'},
    {Name: "ViewAll2", SafeValue: 'ms-Icon--ViewAll2'},
    {Name: "DocumentReply", SafeValue: 'ms-Icon--DocumentReply'},
    {Name: "PlayerSettings", SafeValue: 'ms-Icon--PlayerSettings'},
    {Name: "ReceiptForward", SafeValue: 'ms-Icon--ReceiptForward'},
    {Name: "ReceiptReply", SafeValue: 'ms-Icon--ReceiptReply'},
    {Name: "ReceiptCheck", SafeValue: 'ms-Icon--ReceiptCheck'},
    {Name: "Fax", SafeValue: 'ms-Icon--Fax'},
    {Name: "RecurringEvent", SafeValue: 'ms-Icon--RecurringEvent'},
    {Name: "ReplyAlt", SafeValue: 'ms-Icon--ReplyAlt'},
    {Name: "ReplyAllAlt", SafeValue: 'ms-Icon--ReplyAllAlt'},
    {Name: "EditStyle", SafeValue: 'ms-Icon--EditStyle'},
    {Name: "EditMail", SafeValue: 'ms-Icon--EditMail'},
    {Name: "Lifesaver", SafeValue: 'ms-Icon--Lifesaver'},
    {Name: "LifesaverLock", SafeValue: 'ms-Icon--LifesaverLock'},
    {Name: "InboxCheck", SafeValue: 'ms-Icon--InboxCheck'},
    {Name: "FolderSearch", SafeValue: 'ms-Icon--FolderSearch'},
    {Name: "CollapseMenu", SafeValue: 'ms-Icon--CollapseMenu'},
    {Name: "ExpandMenu", SafeValue: 'ms-Icon--ExpandMenu'},
    {Name: "Boards", SafeValue: 'ms-Icon--Boards'},
    {Name: "SunAdd", SafeValue: 'ms-Icon--SunAdd'},
    {Name: "SunQuestionMark", SafeValue: 'ms-Icon--SunQuestionMark'},
    {Name: "LandscapeOrientation", SafeValue: 'ms-Icon--LandscapeOrientation'},
    {Name: "DocumentSearch", SafeValue: 'ms-Icon--DocumentSearch'},
    {Name: "PublicCalendar", SafeValue: 'ms-Icon--PublicCalendar'},
    {Name: "PublicContactCard", SafeValue: 'ms-Icon--PublicContactCard'},
    {Name: "PublicEmail", SafeValue: 'ms-Icon--PublicEmail'},
    {Name: "PublicFolder", SafeValue: 'ms-Icon--PublicFolder'},
    {Name: "WordDocument", SafeValue: 'ms-Icon--WordDocument'},
    {Name: "PowerPointDocument", SafeValue: 'ms-Icon--PowerPointDocument'},
    {Name: "ExcelDocument", SafeValue: 'ms-Icon--ExcelDocument'},
    {Name: "GroupedList", SafeValue: 'ms-Icon--GroupedList'},
    {Name: "ClassroomLogo", SafeValue: 'ms-Icon--ClassroomLogo'},
    {Name: "Sections", SafeValue: 'ms-Icon--Sections'},
    {Name: "EditPhoto", SafeValue: 'ms-Icon--EditPhoto'},
    {Name: "Starburst", SafeValue: 'ms-Icon--Starburst'},
    {Name: "ShareiOS", SafeValue: 'ms-Icon--ShareiOS'},
    {Name: "AirTickets", SafeValue: 'ms-Icon--AirTickets'},
    {Name: "PencilReply", SafeValue: 'ms-Icon--PencilReply'},
    {Name: "Tiles2", SafeValue: 'ms-Icon--Tiles2'},
    {Name: "SkypeCircleCheck", SafeValue: 'ms-Icon--SkypeCircleCheck'},
    {Name: "SkypeCircleClock", SafeValue: 'ms-Icon--SkypeCircleClock'},
    {Name: "SkypeCircleMinus", SafeValue: 'ms-Icon--SkypeCircleMinus'},
    {Name: "SkypeCheck", SafeValue: 'ms-Icon--SkypeCheck'},
    {Name: "SkypeClock", SafeValue: 'ms-Icon--SkypeClock'},
    {Name: "SkypeMinus", SafeValue: 'ms-Icon--SkypeMinus'},
    {Name: "SkypeMessage", SafeValue: 'ms-Icon--SkypeMessage'},
    {Name: "ClosedCaption", SafeValue: 'ms-Icon--ClosedCaption'},
    {Name: "ATPLogo", SafeValue: 'ms-Icon--ATPLogo'},
    {Name: "OfficeFormLogo", SafeValue: 'ms-Icon--OfficeFormLogo'},
    {Name: "RecycleBin", SafeValue: 'ms-Icon--RecycleBin'},
    {Name: "EmptyRecycleBin", SafeValue: 'ms-Icon--EmptyRecycleBin'},
    {Name: "Hide2", SafeValue: 'ms-Icon--Hide2'},
    {Name: "iOSAppStoreLogo", SafeValue: 'ms-Icon--iOSAppStoreLogo'},
    {Name: "AndroidLogo", SafeValue: 'ms-Icon--AndroidLogo'},
    {Name: "Breadcrumb", SafeValue: 'ms-Icon--Breadcrumb'},
    {Name: "ClearFilter", SafeValue: 'ms-Icon--ClearFilter'},
    {Name: "Flow", SafeValue: 'ms-Icon--Flow'},
    {Name: "PowerAppsLogo", SafeValue: 'ms-Icon--PowerAppsLogo'},
    {Name: "PowerApps2Logo", SafeValue: 'ms-Icon--PowerApps2Logo'}
  ];

  private latestValidateValue: string;
  private async: Async;
  private delayedValidate: (value: string) => void;
  private _key: string;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldIconPickerHostProps) {
    super(props);

    //Bind the current object to the external called onSelectDate method
    this.onOpenDialog = this.onOpenDialog.bind(this);
    this.toggleHover = this.toggleHover.bind(this);
    this.toggleHoverLeave = this.toggleHoverLeave.bind(this);
    this.onClickFont = this.onClickFont.bind(this);
    this.onFontDropdownChanged = this.onFontDropdownChanged.bind(this);
    this.mouseEnterDropDown = this.mouseEnterDropDown.bind(this);
    this.mouseLeaveDropDown = this.mouseLeaveDropDown.bind(this);
    this._key = GuidHelper.getGuid();

    if (this.props.orderAlphabetical === true)
      this.orderAlphabetical();

    //Init the state
    this.state = {
        isOpen: false,
        isHoverDropdown: false,
        errorMessage: ''
      };

    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);

    //Inits the default value
    if (props.initialValue != null && props.initialValue != '') {
      for (var i = 0; i < this.fonts.length; i++) {
        var font = this.fonts[i];
        if (font.SafeValue == props.initialValue) {
          this.state.selectedFont = font.Name;
          this.state.safeSelectedFont = font.SafeValue;
        }
      }
    }
  }

  /**
   * @function
   * Orders the font list
   */
  private orderAlphabetical(): void {
    this.fonts.sort(this.compare);
  }

  private compare(a: ISafeFont, b: ISafeFont) {
    if (a.Name < b.Name)
      return -1;
    if (a.Name > b.Name)
      return 1;
    return 0;
  }

  /**
   * @function
   * Function to refresh the Web Part properties
   */
  private changeSelectedFont(newValue: string): void {
    this.delayedValidate(newValue);
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: string): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.initialValue, value);
      return;
    }

    if (this.latestValidateValue === value)
      return;
    this.latestValidateValue = value;

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || '');
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.initialValue, value);
        this.state.errorMessage = result;
        this.setState(this.state);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.initialValue, value);
          this.state.errorMessage = errorMessage;
          this.setState(this.state);
        });
      }
    }
    else {
      this.notifyAfterValidate(this.props.initialValue, value);
    }
  }

  /**
   * @function
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: string, newValue: string) {
    if (this.props.onPropertyChange && newValue != null) {
      this.props.properties[this.props.targetProperty] = newValue;
      this.props.onPropertyChange(this.props.targetProperty, oldValue, newValue);
    }
  }

  /**
   * @function
   * Called when the component will unmount
   */
  public componentWillUnmount() {
    this.async.dispose();
  }

  /**
   * @function
   * Function to open the dialog
   */
  private onOpenDialog(): void {
    if (this.props.disabled === true)
      return;
    this.state.isOpen = !this.state.isOpen;
    this.setState(this.state);
  }

  /**
   * @function
   * Mouse is hover a font
   */
  private toggleHover(element?: any) {
    var hoverFont: string = element.currentTarget.textContent;
    this.state.hoverFont = hoverFont;
    this.setState(this.state);
  }

  /**
   * @function
   * Mouse is leaving a font
   */
  private toggleHoverLeave(element?: any) {
    this.state.hoverFont = '';
    this.setState(this.state);
  }

  /**
   * @function
   * Mouse is hover the fontpicker
   */
  private mouseEnterDropDown(element?: any) {
    this.state.isHoverDropdown = true;
    this.setState(this.state);
  }

  /**
   * @function
   * Mouse is leaving the fontpicker
   */
  private mouseLeaveDropDown(element?: any) {
    this.state.isHoverDropdown = false;
    this.setState(this.state);
  }

  /**
   * @function
   * User clicked on a font
   */
  private onClickFont(element?: any) {
    var clickedFont: string = element.currentTarget.textContent;
    this.state.selectedFont = clickedFont;
    this.state.safeSelectedFont = this.getSafeFont(clickedFont);
    this.onOpenDialog();
    this.changeSelectedFont(this.state.safeSelectedFont);
    this.setState(this.state);
  }

  /**
   * @function
   * Gets a safe font value from a font name
   */
  private getSafeFont(fontName: string): string {
    for (var i = 0; i < this.fonts.length; i++) {
      var font = this.fonts[i];
      if (font.Name === fontName)
        return font.SafeValue;
    }
    return '';
  }

  /**
   * @function
   * The font dropdown selected value changed (used when the previewFont property equals false)
   */
  private onFontDropdownChanged(option: IDropdownOption, index?: number): void {
    this.changeSelectedFont(option.key as string);
  }

  /**
   * @function
   * Renders the controls
   */
  public render(): JSX.Element {

    if (this.props.preview === false) {
      //If the user don't want to use the preview font picker,
      //we're building a classical drop down picker
      var dropDownOptions: IDropdownOption[] = [];
      var selectedKey: string;
      this.fonts.map((font: ISafeFont) => {
        var isSelected: boolean = false;
        isSelected = true;
        selectedKey = font.SafeValue;
        dropDownOptions.push(
          {
            key: font.SafeValue,
            text: font.Name,
            isSelected: isSelected
          }
        );
      });
      return (
        <div>
          <Dropdown label={this.props.label} options={dropDownOptions} selectedKey={selectedKey}
            onChanged={this.onFontDropdownChanged} disabled={this.props.disabled} />
          { this.state.errorMessage != null && this.state.errorMessage != '' && this.state.errorMessage != undefined ?
              <div><div aria-live='assertive' className='ms-u-screenReaderOnly' data-automation-id='error-message'>{  this.state.errorMessage }</div>
              <span>
                <p className='ms-TextField-errorMessage ms-u-slideDownIn20'>{ this.state.errorMessage }</p>
              </span>
              </div>
            : ''}
        </div>
      );
    }
    else {
      //User wants to use the preview font picker, so just build it
      var fontSelect = {
        fontSize: '16px',
        width: '100%',
        position: 'relative',
        display: 'inline-block',
        zoom: 1
      };
      var dropdownColor = '1px solid #c8c8c8';
      if (this.props.disabled === true)
        dropdownColor = '1px solid #f4f4f4';
      else if (this.state.isOpen === true)
        dropdownColor = '1px solid #3091DE';
      else if (this.state.isHoverDropdown === true)
        dropdownColor = '1px solid #767676';
      var fontSelectA = {
        backgroundColor: this.props.disabled === true ? '#f4f4f4' : '#fff',
        borderRadius        : '0px',
        backgroundClip        : 'padding-box',
        border: dropdownColor,
        display: 'block',
        overflow: 'hidden',
        whiteSpace: 'nowrap',
        position: 'relative',
        height: '26px',
        lineHeight: '26px',
        padding: '0 0 0 8px',
        color: this.props.disabled === true ? '#a6a6a6' : '#444',
        textDecoration: 'none',
        cursor: this.props.disabled === true ? 'default' : 'pointer'
      };
      var fontSelectASpan = {
        marginRight: '26px',
        display: 'block',
        overflow: 'hidden',
        whiteSpace: 'nowrap',
        lineHeight: '1.8',
        textOverflow: 'ellipsis',
        cursor: this.props.disabled === true ? 'default' : 'pointer',
        //fontFamily: this.state.safeSelectedFont != null && this.state.safeSelectedFont != '' ? this.state.safeSelectedFont : 'Arial',
        //fontSize: this.state.safeSelectedFont,
        fontWeight: 400
      };
      var fontSelectADiv = {
        borderRadius        : '0 0px 0px 0',
        backgroundClip        : 'padding-box',
        border: '0px',
        position: 'absolute',
        right: '0',
        top: '0',
        display: 'block',
        height: '100%',
        width: '22px'
      };
      var fontSelectADivB = {
        display: 'block',
        width: '100%',
        height: '100%',
        cursor: this.props.disabled === true ? 'default' : 'pointer',
        marginTop: '2px'
      };
      var fsDrop = {
        background: '#fff',
        border: '1px solid #aaa',
        borderTop: '0',
        position: 'absolute',
        top: '29px',
        left: '0',
        width: 'calc(100% - 2px)',
        //boxShadow: '0 4px 5px rgba(0,0,0,.15)',
        zIndex: 999,
        display: this.state.isOpen ? 'block' : 'none'
      };
      var fsResults = {
        margin: '0 4px 4px 0',
        maxHeight: '190px',
        width: 'calc(100% - 4px)',
        padding: '0 0 0 4px',
        position: 'relative',
        overflowX: 'hidden',
        overflowY: 'auto'
      };
      var carret: string = this.state.isOpen ? 'ms-Icon ms-Icon--ChevronUp' : 'ms-Icon ms-Icon--ChevronDown';
      //Renders content
      return (
        <div style={{ marginBottom: '8px'}}>
          <Label>{this.props.label}</Label>
          <div style={fontSelect}>
            <a style={fontSelectA} onClick={this.onOpenDialog}
              onMouseEnter={this.mouseEnterDropDown} onMouseLeave={this.mouseLeaveDropDown} role="menuitem">
              <span style={fontSelectASpan}>
                <i className={'ms-Icon ms-Icon--' + this.state.selectedFont} aria-hidden="true" style={{marginRight:'10px'}}></i>
                {this.state.selectedFont}
              </span>
              <div style={fontSelectADiv}>
                <i style={fontSelectADivB} className={carret}></i>
              </div>
            </a>
            <div style={fsDrop}>
              <ul style={fsResults}>
                {this.fonts.map((font: ISafeFont, index: number) => {
                  var backgroundColor: string = 'transparent';
                  if (this.state.selectedFont === font.Name)
                    backgroundColor = '#c7e0f4';
                  else if (this.state.hoverFont === font.Name)
                    backgroundColor = '#eaeaea';
                  var innerStyle = {
                    lineHeight: '80%',
                    padding: '7px 7px 8px',
                    margin: '0',
                    listStyle: 'none',
                    fontSize: '16px',
                    backgroundColor: backgroundColor,
                    cursor: 'pointer'
                  };
                  return (
                    <li value={font.Name}  role="menuitem"
                      key={this._key + '-iconpicker-' + index}
                      onMouseEnter={this.toggleHover} onClick={this.onClickFont}
                      onMouseLeave={this.toggleHoverLeave} style={innerStyle}>
                      <i className={'ms-Icon ' + font.SafeValue} aria-hidden="true" style={{fontSize: '24px', marginRight:'10px'}}></i>
                      {font.Name}
                    </li>
                  );
                })
                }
              </ul>
            </div>
          </div>
          { this.state.errorMessage != null && this.state.errorMessage != '' && this.state.errorMessage != undefined ?
              <div><div aria-live='assertive' className='ms-u-screenReaderOnly' data-automation-id='error-message'>{  this.state.errorMessage }</div>
              <span>
                <p className='ms-TextField-errorMessage ms-u-slideDownIn20'>{ this.state.errorMessage }</p>
              </span>
              </div>
            : ''}
        </div>
      );
    }
  }
}