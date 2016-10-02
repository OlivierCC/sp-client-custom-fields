declare interface IStrings {

  //SPListQuery labels
  //Copy the following labels in your project if you want to use the SPListQuery
  SPListQueryList: string;
  SPListQueryOrderBy: string;
  SPListQueryArranged: string;
  SPListQueryMax: string;
  SPListQueryAdd: string;
  SPListQueryRemove: string;
  SPListQueryOperatorEq: string;
  SPListQueryOperatorNe: string;
  SPListQueryOperatorStartsWith: string;
  SPListQueryOperatorSubstringof: string;
  SPListQueryOperatorLt: string;
  SPListQueryOperatorLe: string;
  SPListQueryOperatorGt: string;
  SPListQueryOperatorGe: string;

  //PicturePicker labels
  //Copy the following labels in your project if you want to use the PicturePicker
  PicturePickerTitle: string;
  PicturePickerRecent: string;
  PicturePickerSite: string;
  PicturePickerButtonSelect: string;
  PicturePickerButtonReset: string;

  //DocumentPicker labels
  //Copy the following labels in your project if you want to use the DocumentPicker
  DocumentPickerTitle: string;
  DocumentPickerRecent: string;
  DocumentPickerSite: string;
  DocumentPickerButtonSelect: string;
  DocumentPickerButtonReset: string;

  //PeoplePicker labels
  //Copy the following labels in your project if you want to use the PeoplePicker
  PeoplePickerSuggestedContacts: string;
  PeoplePickerNoResults: string;
  PeoplePickerLoading: string;

  //SPListPicker labels
  //Copy the following labels in your project if you want to use the SPListPicker
  SPListPickerLoading: string;

  //SPFolderPicker labels
  //Copy the following labels in your project if you want to use the SPFolderPicker
  SPFolderPickerDialogTitle: string;
  SPFolderPickerSelectButton: string;
  SPFolderPickerCancelButton: string;

  //DatePicker labels
  //Copy the following labels in your project if you want to use the DatePicker
  DatePickerMonthLongJanuary: string;
  DatePickerMonthShortJanuary: string;
  DatePickerMonthLongFebruary: string;
  DatePickerMonthShortFebruary: string;
  DatePickerMonthLongMarch: string;
  DatePickerMonthShortMarch: string;
  DatePickerMonthLongApril: string;
  DatePickerMonthShortApril: string;
  DatePickerMonthLongMay: string;
  DatePickerMonthShortMay: string;
  DatePickerMonthLongJune: string;
  DatePickerMonthShortJune: string;
  DatePickerMonthLongJuly: string;
  DatePickerMonthShortJuly: string;
  DatePickerMonthLongAugust: string;
  DatePickerMonthShortAugust: string;
  DatePickerMonthLongSeptember: string;
  DatePickerMonthShortSeptember: string;
  DatePickerMonthLongOctober: string;
  DatePickerMonthShortOctober: string;
  DatePickerMonthLongNovember: string;
  DatePickerMonthShortNovember: string;
  DatePickerMonthLongDecember: string;
  DatePickerMonthShortDecember: string;
  DatePickerDayLongSunday: string;
  DatePickerDayShortSunday: string;
  DatePickerDayLongMonday: string;
  DatePickerDayShortMonday: string;
  DatePickerDayLongTuesday: string;
  DatePickerDayShortTuesday: string;
  DatePickerDayLongWednesday: string;
  DatePickerDayShortWednesday: string;
  DatePickerDayLongThursday: string;
  DatePickerDayShortThursday: string;
  DatePickerDayLongFriday: string;
  DatePickerDayShortFriday: string;
  DatePickerDayLongSaturday: string;
  DatePickerDayShortSaturday: string;
}

declare module 'propertyFieldsStrings' {
  const strings: IStrings;
  export = strings;
}
