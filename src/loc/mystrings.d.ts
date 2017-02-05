declare interface IStrings {

  //DimensionPicker labels
  DimensionWidth: string;
  DimensionHeight: string;
  DimensionRatio: string;

  //CustomList labels
  CustomListAddItem: string;
  CustomListBack: string;
  CustomListTrue: string;
  CustomListFalse: string;
  CustomListOK: string;
  CustomListCancel: string;
  CustomListEdit: string;
  CustomListDel: string;
  CustomListYes: string;
  CustomListNo: string;
  CustomListConfirmDel: string;
  CustomListConfirmDelMssg: string;
  CustomListFieldMissing: string;

  //SPListQuery labels
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
  PicturePickerTitle: string;
  PicturePickerRecent: string;
  PicturePickerSite: string;
  PicturePickerButtonSelect: string;
  PicturePickerButtonReset: string;

  //DocumentPicker labels
  DocumentPickerTitle: string;
  DocumentPickerRecent: string;
  DocumentPickerSite: string;
  DocumentPickerButtonSelect: string;
  DocumentPickerButtonReset: string;

  //PeoplePicker labels
  PeoplePickerSuggestedContacts: string;
  PeoplePickerNoResults: string;
  PeoplePickerLoading: string;

  //SPListPicker labels
  SPListPickerLoading: string;

  //SPFolderPicker labels
  SPFolderPickerDialogTitle: string;
  SPFolderPickerSelectButton: string;
  SPFolderPickerCancelButton: string;

  //DatePicker labels
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

declare module 'sp-client-custom-fields/strings' {
  const strings: IStrings;
  export = strings;
}
