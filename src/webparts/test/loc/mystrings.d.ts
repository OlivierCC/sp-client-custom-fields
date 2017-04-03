declare interface ITestStrings {
  //Web Parts property labels
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  DateFieldLabel: string;
  DateTimeFieldLabel: string;
  PeopleFieldLabel: string;
  GroupFieldLabel: string;
  SPListFieldLabel: string;
  ColorFieldLabel: string;
  ColorMiniFieldLabel: string;
  SPFolderFieldLabel: string;
  PasswordFieldLabel: string;
  FontFieldLabel: string;
  FontSizeFieldLabel: string;
  PhoneNumberFieldLabel: string;
  MaskedInputFieldLabel: string;
  GeoLocationFieldLabel: string;
  PictureFieldLabel: string;
  IconFieldLabel: string;
  DocumentFieldLabel: string;
  DisplayModeFieldLabel: string;
  CustomListFieldLabel: string;
  QueryFieldLabel: string;
  AlignFieldLabel: string;
  DropDownSelectFieldLabel: string;
  RichTextBoxFieldLabel: string;
  SliderRangeFieldLabel: string;
  DimensionFieldLabel: string;
  SortableListFieldLabel: string;
  TreeViewFieldLabel: string;
  DropDownTreeViewFieldLabel: string;
  TagPickerFieldLabel: string;
  StarRatingFieldLabel: string;
  NumericInputFieldLabel: string;
  AutoSuggestFieldLabel: string;
  SearchPropertiesFieldLabel: string;
}

declare module 'testStrings' {
  const strings: ITestStrings;
  export = strings;
}
