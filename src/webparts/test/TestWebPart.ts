/**
 * @file TestWebPart.ts
 * Custom field implementation sample for the SharePoint Framework (SPfx)
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IWebPartContext
} from '@microsoft/sp-webpart-base';
import { Version } from '@microsoft/sp-core-library';

import * as strings from 'testStrings';
import Test, { ITestProps } from './components/Test';
import { ITestWebPartProps } from './ITestWebPartProps';

//Include the PropertyFieldDatePicker component
import { PropertyFieldDatePicker } from '../../PropertyFieldDatePicker';
//Include the PropertyFieldDateTimePicker component
import { PropertyFieldDateTimePicker, ITimeConvention } from '../../PropertyFieldDateTimePicker';
//Include the PropertyFieldColorPicker component
import { PropertyFieldColorPicker } from '../../PropertyFieldColorPicker';
//Include the PropertyFieldColorPickerMini component
import { PropertyFieldColorPickerMini } from '../../PropertyFieldColorPickerMini';
//Include the PropertyFieldPeoplePicker component
import { PropertyFieldPeoplePicker } from '../../PropertyFieldPeoplePicker';
//Include the PropertyFieldSPListPicker component
import { PropertyFieldSPListPicker, PropertyFieldSPListPickerOrderBy } from '../../PropertyFieldSPListPicker';
//Include the PropertyFieldSPListMultiplePicker component
import { PropertyFieldSPListMultiplePicker, PropertyFieldSPListMultiplePickerOrderBy } from '../../PropertyFieldSPListMultiplePicker';
//Include the PropertyFieldSPFolderPicker component
import { PropertyFieldSPFolderPicker } from '../../PropertyFieldSPFolderPicker';
//Include the PropertyFieldPassword component
import { PropertyFieldPassword } from '../../PropertyFieldPassword';
//Include the PropertyFieldFontPicker component
import { PropertyFieldFontPicker } from '../../PropertyFieldFontPicker';
//Include the PropertyFieldFontSizePicker component
import { PropertyFieldFontSizePicker } from '../../PropertyFieldFontSizePicker';
//Include the PropertyFieldPhoneNumber component
import { PropertyFieldPhoneNumber, IPhoneNumberFormat } from '../../PropertyFieldPhoneNumber';
//Include the PropertyFieldMaskedInput component
import { PropertyFieldMaskedInput } from '../../PropertyFieldMaskedInput';
//Include the PropertyFieldMaskedInput component
import { PropertyFieldMapPicker } from '../../PropertyFieldMapPicker';
//Include the PropertyFieldPicturePicker component
import { PropertyFieldPicturePicker } from '../../PropertyFieldPicturePicker';
//Include the PropertyFieldIconPicker component
import { PropertyFieldIconPicker } from '../../PropertyFieldIconPicker';
//Include the PropertyFieldDocumentPicker component
import { PropertyFieldDocumentPicker } from '../../PropertyFieldDocumentPicker';
//Include the PropertyFieldDisplayMode component
import { PropertyFieldDisplayMode } from '../../PropertyFieldDisplayMode';
//Include the PropertyFieldCustomList component
import { PropertyFieldCustomList, CustomListFieldType } from '../../PropertyFieldCustomList';
//Include the PropertyFieldSPListQuery component
import { PropertyFieldSPListQuery, PropertyFieldSPListQueryOrderBy } from '../../PropertyFieldSPListQuery';
//Include the PropertyFieldAlignPicker component
import { PropertyFieldAlignPicker } from '../../PropertyFieldAlignPicker';
//Include the PropertyFieldDropDownSelect component
import { PropertyFieldDropDownSelect } from '../../PropertyFieldDropDownSelect';
//Include the PropertyFieldRichTextBox component
import { PropertyFieldRichTextBox } from '../../PropertyFieldRichTextBox';
//Include the PropertyFieldSliderRange component
import { PropertyFieldSliderRange } from '../../PropertyFieldSliderRange';
//Include the PropertyFieldDimensionPicker component
import { PropertyFieldDimensionPicker } from '../../PropertyFieldDimensionPicker';
//Include the PropertyFieldSortableList component
import { PropertyFieldSortableList, ISortableListOrder } from '../../PropertyFieldSortableList';
//Include the PropertyFieldTreeView component
import { PropertyFieldTreeView } from '../../PropertyFieldTreeView';
//Include the PropertyFieldDropDownTreeView component
import { PropertyFieldDropDownTreeView } from '../../PropertyFieldDropDownTreeView';
//Include the PropertyFieldTagPicker component
import { PropertyFieldTagPicker } from '../../PropertyFieldTagPicker';
//Include the PropertyFieldStarRating component
import { PropertyFieldStarRating } from '../../PropertyFieldStarRating';
//Include the PropertyFieldGroupPicker component
import { PropertyFieldGroupPicker, IGroupType } from '../../PropertyFieldGroupPicker';
//Include the PropertyFieldNumericInput component
import { PropertyFieldNumericInput } from '../../PropertyFieldNumericInput';

export default class TestWebPart extends BaseClientSideWebPart<ITestWebPartProps> {

  public constructor(context: IWebPartContext) {
    super();

    //Hack: to invoke correctly the onPropertyChange function outside this class
    //we need to bind this object on it first
    this.onPropertyPaneFieldChanged = this.onPropertyPaneFieldChanged.bind(this);
    this.testPropertyChanged = this.testPropertyChanged.bind(this);

  }

  public render(): void {
    const element: React.ReactElement<ITestProps> = React.createElement(Test, {
      description: this.properties.description,
      color: this.properties.color,
      miniColor: this.properties.miniColor,
      date: this.properties.date,
      date2: this.properties.date2,
      datetime: this.properties.datetime,
      folder: this.properties.folder,
      people: this.properties.people,
      groups: this.properties.groups,
      list: this.properties.list,
      listsCollection: this.properties.listsCollection,
      password: this.properties.password,
      numeric: this.properties.numeric,
      font: this.properties.font,
      fontSize: this.properties.fontSize,
      phone: this.properties.phone,
      maskedInput: this.properties.maskedInput,
      geolocation: this.properties.geolocation,
      picture: this.properties.picture,
      icon: this.properties.icon,
      document: this.properties.document,
      displayMode: this.properties.displayMode,
      customList: this.properties.customList,
      query: this.properties.query,
      align: this.properties.align,
      richTextBox: this.properties.richTextBox,
      dropDownSelect: this.properties.dropDownSelect,
      sliderRange: this.properties.sliderRange,
      dimension: this.properties.dimension,
      sortableList: this.properties.sortableList,
      treeView: this.properties.treeView,
      dropDownTreeView: this.properties.dropDownTreeView,
      tags: this.properties.tags,
      starRating: this.properties.starRating
    });

    ReactDom.render(element, this.domElement);
  }

	protected get disableReactivePropertyChanges(): boolean {
		return false;
	}

  private formatDateIso(date: Date): string {
    //example for ISO date formatting
    return date.toISOString();
  }

  private testPropertyChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this.properties.font = newValue;
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }


/*
  //Samples of validation method, to use
  //with the onGetErrorMessage function of Propery Fields.
  //See https://dev.office.com/sharepoint/docs/spfx/web-parts/guidance/validate-web-part-property-values

  private canNotBeEmpty(value: string): string {
    if (value === null || value.trim().length === 0) {
      return 'Provide a value';
    }
    return '';
  }

  private canNotBeEmptyPromise(value: string): Promise<string> {
    return new Promise<string>((resolve: (validationErrorMessage: string) => void, reject: (error: any) => void): void => {
      if (value === null || value.length === 0) {
        resolve('Provide a value');
        return;
      }
      resolve('');
    });
  }

  private canNotBeArial(value: string): string {
    if (value === null || value.trim().length === 0) {
      return '';
    }
    if (value.indexOf("Arial") !== -1)
      return 'Font can not be Arial';
    return '';
  }

  private canNotBeXSmall(value: string): string {
    if (value === null || value.trim().length === 0) {
      return '';
    }
    if (value.indexOf("x-small") !== -1)
      return 'Font size can not be x-small';
    return '';
  }

  private canNotBeAADLogo(value: string): string {
    if (value === null || value.trim().length === 0) {
      return '';
    }
    if (value.indexOf("AADLogo") !== -1)
      return 'Icon can not be AADLogo';
    return '';
  }

  private canNotBeBlack(value: string): string {
    if (value === null || value.trim().length === 0) {
      return '';
    }
    if (value.indexOf("#000000") !== -1)
      return 'Color can not be black';
    return '';
  }

  private canNotBeAlignLeft(value: string): string {
    if (value === null || value.trim().length === 0) {
      return '';
    }
    if (value.indexOf("left") !== -1)
      return 'Align can not be left';
    return '';
  }

  private arrayCanNotBeEmpty(value: string[]): string {
    if (value === null || value.length === 0) {
      return 'Array can not be empty';
    }
    return '';
  }

  private canNotBeIn2016(value: string): string {
    if (value === null || value.trim().length === 0) {
      return '';
    }
    if (value.indexOf("2016") !== -1)
      return 'Date can not be during 2016.';
    return '';
  }

  private canNotBe0Location(value: string): string {
    if (value === null || value.trim().length === 0) {
      return '';
    }
    if (value == '0,0')
      return 'Bad geoLocation.';
    return '';
  }

  private badPhoneNumber(value: string): string {
    if (value === null || value.trim().length === 0 || value == '(') {
      return 'Provide a valid phone number.';
    }
    if (value.indexOf("(00") != 0) {
      return 'Phone number must be begin with (00.';
    }
    return '';
  }

  private canNotSelectThisList(value: string): string {
    if (value === null || value.trim().length === 0) {
      return '';
    }
    if (value.indexOf("6770c83b") !== -1) {
      return 'You can not select this list.';
    }
    return '';
  }

  private canNotBeMock(value: string): string {
    if (value === null || value.trim().length === 0) {
      return '';
    }
    if (value.indexOf('mock'))
      return 'The mock is not allowed';
    return '';
  }

  private canNotBeList(value: string): string {
    if (value === null || value.trim().length === 0) {
      return '';
    }
    if (value.indexOf("list") !== -1) {
      return 'You can not select the list mode.';
    }
    return '';
  }

  private canNotBeOrderById(value: string): string {
    if (value === null || value.trim().length === 0) {
      return '';
    }
    if (value.indexOf("orderBy=ID") !== -1) {
      return 'You can not order by ID.';
    }
    return '';
  }

  private canNotChooseMoreThan2People(value: any[]): string {
    if (value.length > 2) {
      return 'You can not choose more than 2 people.';
    }
    return '';
  }

  private canNotBeDoc(value: string): string {
    if (value === null || value.trim().length === 0) {
      return '';
    }
    if (value.indexOf(".doc") !== -1) {
      return 'You can not choose a *.doc file.';
    }
    return '';
  }

  private invalidRange(value: string): string {
    if (value === null || value.trim().length === 0) {
      return '';
    }
    if (value === '0,500') {
      return 'Invalid range.';
    }
    return '';
  }
*/

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          //Display the web part properties as accordion
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: 'Layout Fields',
              groupFields: [
                PropertyFieldFontPicker('font', {
                  label: strings.FontFieldLabel,
                  useSafeFont: true,
                  previewFonts: true,
                  initialValue: this.properties.font,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'fontFieldId'
                }),
                PropertyFieldFontSizePicker('fontSize', {
                  label: strings.FontSizeFieldLabel,
                  usePixels: false,
                  preview: true,
                  initialValue: this.properties.fontSize,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'fontSizeFieldId'
                }),
                PropertyFieldFontSizePicker('fontSize', {
                  label: strings.FontSizeFieldLabel,
                  usePixels: true,
                  preview: true,
                  initialValue: this.properties.fontSize,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'fontSizeField2Id'
                }),
                PropertyFieldIconPicker('icon', {
                  label: strings.IconFieldLabel,
                  initialValue: this.properties.icon,
                  orderAlphabetical: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'iconFieldId'
                }),
                PropertyFieldColorPickerMini('miniColor', {
                  label: strings.ColorMiniFieldLabel,
                  initialColor: this.properties.miniColor,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'colorMiniFieldId'
                }),
                PropertyFieldColorPicker('color', {
                  label: strings.ColorFieldLabel,
                  initialColor: this.properties.color,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'colorFieldId'
                }),
                PropertyFieldAlignPicker('align', {
                  label: strings.AlignFieldLabel,
                  initialValue: this.properties.align,
                  onPropertyChanged: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'alignFieldId'
                }),
                PropertyFieldDimensionPicker('dimension', {
                  label: strings.DimensionFieldLabel,
                  initialValue: this.properties.dimension,
                  preserveRatio: true,
                  preserveRatioEnabled: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'dimensionFieldId'
                })
              ],
            },
            {
              groupName: 'Text Input Fields',
              groupFields: [

                PropertyFieldCustomList('customList', {
                  label: strings.CustomListFieldLabel,
                  value: this.properties.customList,
                  headerText: "Manage News",
                  fields: [
                    { title: 'News Title', required: true, type: CustomListFieldType.string },
                    { title: 'Sub title', required: true, type: CustomListFieldType.string },
                    { title: 'Link', required: false, type: CustomListFieldType.string, hidden: true },
                    { title: 'Order', required: true, type: CustomListFieldType.number },
                    { title: 'Active', required: false, type: CustomListFieldType.boolean },
                    { title: 'Start Date', required: false, type: CustomListFieldType.date, hidden: true },
                    { title: 'End Date', required: false, type: CustomListFieldType.date, hidden: true },
                    { title: 'Picture', required: false, type: CustomListFieldType.picture, hidden: true },
                    { title: 'Users', required: false, type: CustomListFieldType.users, hidden: true }
                    /*,
                    { title: 'Font', required: false, type: CustomListFieldType.font, hidden: true },
                    { title: 'Font size', required: false, type: CustomListFieldType.fontSize, hidden: true },
                    { title: 'Icon', required: false, type: CustomListFieldType.icon, hidden: true },
                    { title: 'Password', required: false, type: CustomListFieldType.password, hidden: true },
                    { title: 'Users', required: false, type: CustomListFieldType.users, hidden: true },
                    { title: 'Color', required: false, type: CustomListFieldType.color, hidden: true },
                    { title: 'List', required: false, type: CustomListFieldType.list, hidden: true },
                    { title: 'Picture', required: false, type: CustomListFieldType.picture, hidden: true },
                    { title: 'Document', required: false, type: CustomListFieldType.document, hidden: true },
                    { title: 'Folder', required: false, type: CustomListFieldType.folder, hidden: true }
                    */
                  ],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  disabled: false,
                  key: 'customListFieldId'
                }),
                PropertyFieldDropDownSelect('dropDownSelect', {
                  label: strings.DropDownSelectFieldLabel,
                  options: [
                    {'key': 'Option1', 'text': 'Option 1'},
                    {'key': 'Option2', 'text': 'Option 2'},
                    {'key': 'Option3', 'text': 'Option 3'},
                    {'key': 'Option4', 'text': 'Option 4'},
                    {'key': 'Option5', 'text': 'Option 5'},
                    {'key': 'Option6', 'text': 'Option 6'},
                    {'key': 'Option7', 'text': 'Option 7'}
                  ],
                  initialValue: this.properties.dropDownSelect,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'dropDownSelectFieldId'
                }),
                PropertyFieldSortableList('sortableList', {
                  label: strings.SortableListFieldLabel,
                  items: [
                    {'key': 'Option1', 'text': 'Option 1'},
                    {'key': 'Option2', 'text': 'Option 2'},
                    {'key': 'Option3', 'text': 'Option 3'},
                    {'key': 'Option4', 'text': 'Option 4'},
                    {'key': 'Option5', 'text': 'Option 5'}
                  ],
                  selectedItems: this.properties.sortableList,
                  sortBy: ISortableListOrder.Text,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'sortableListFieldId'
                }),
                PropertyFieldDropDownTreeView('dropDownTreeView', {
                  label: strings.DropDownTreeViewFieldLabel,
                  tree: [
                    {
                      id: 'Analytics', label: 'Analytics',
                      children: [
                        {
                          id: 'Market analyses', label: 'Market analyses',
                          collapsed: true,
                          children: [{
                            id: 'Key-on-screen.jpg', label: 'Key-on-screen.jpg',
                            leaf: true
                          }]
                        },
                        {
                          id: 'Northwind marketing', label: 'Northwind marketing',
                          children: [{
                            id: 'New Product Overview.pptx',
                            label: 'New Product Overview.pptx',
                            leaf: true
                          }, {
                            id: 'RD Expenses Q1 to Q3.xlsx', label: 'RD Expenses Q1 to Q3.xlsx',
                            leaf: true
                          }, {
                            id: 'Sat Survey.xlsx', label: 'Sat Survey.xlsx',
                            leaf: true
                          }]
                        },
                        {
                          id: 'Project Budget Audit.docx', label: 'Project Budget Audit.docx',
                          leaf: true
                        }, {
                          id: 'Engineering Costs Q1.pptx', label: 'Engineering Costs Q1.pptx',
                          leaf: true
                        }]
                    },
                    {
                      id: 'Notebooks', label: 'Notebooks',
                      children: [{
                        id: 'New Project Timeline.docx', label: 'New Project Timeline.docx',
                        leaf: true
                      }, {
                        id: 'Marketing Video.mp4', label: 'Marketing Video.mp4',
                        leaf: true
                      }, {
                        id: 'Meeting Audio Record.mp3', label: 'Meeting Audio Record.mp3',
                        leaf: true
                      }]
                    }
                  ],
                  selectedNodesIDs: this.properties.dropDownTreeView,
                  allowMultipleSelections: true,
                  allowFoldersSelections: false,
                  nodesPaddingLeft: 15,
                  checkboxEnabled: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'dropDownTreeViewFieldId'
                }),
                PropertyFieldTreeView('treeView', {
                  label: strings.TreeViewFieldLabel,
                  tree: [
                    {
                      id: 'Analytics', label: 'Analytics',
                      pictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2012/96/iconmonstr-folder-1.png',
                      expandedPictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2012/96/iconmonstr-folder-20.png',
                      children: [
                        {
                          id: 'Market analyses', label: 'Market analyses',
                          collapsed: true,
                          pictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2012/96/iconmonstr-folder-1.png',
                          expandedPictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2012/96/iconmonstr-folder-20.png',
                          children: [{
                            id: 'Key-on-screen.jpg', label: 'Key-on-screen.jpg',
                            pictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2012/96/iconmonstr-picture-1.png',
                            leaf: true
                          }]
                        },
                        {
                          id: 'Northwind marketing', label: 'Northwind marketing',
                          pictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2012/96/iconmonstr-folder-1.png',
                          expandedPictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2012/96/iconmonstr-folder-20.png',
                          children: [{
                            id: 'New Product Overview.pptx',
                            label: 'New Product Overview.pptx',
                            pictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2014/96/iconmonstr-flip-chart-2.png',
                            leaf: true
                          }, {
                            id: 'RD Expenses Q1 to Q3.xlsx', label: 'RD Expenses Q1 to Q3.xlsx',
                            pictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2017/96/iconmonstr-flip-chart-9.png',
                            leaf: true
                          }, {
                            id: 'Sat Survey.xlsx', label: 'Sat Survey.xlsx',
                            pictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2017/96/iconmonstr-flip-chart-9.png',
                            leaf: true
                          }]
                        },
                        {
                          id: 'Project Budget Audit.docx', label: 'Project Budget Audit.docx',
                          pictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2013/96/iconmonstr-note-14.png',
                          selectedPictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2013/96/iconmonstr-note-13.png',
                          leaf: true
                        }, {
                          id: 'Engineering Costs Q1.pptx', label: 'Engineering Costs Q1.pptx',
                          pictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2014/96/iconmonstr-flip-chart-2.png',
                          leaf: true
                        }]
                    },
                    {
                      id: 'Notebooks', label: 'Notebooks',
                      pictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2012/96/iconmonstr-folder-1.png',
                      expandedPictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2012/96/iconmonstr-folder-20.png',
                      children: [{
                        id: 'New Project Timeline.docx', label: 'New Project Timeline.docx',
                        pictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2013/96/iconmonstr-note-14.png',
                        selectedPictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2013/96/iconmonstr-note-13.png',
                        leaf: true
                      }, {
                        id: 'Marketing Video.mp4', label: 'Marketing Video.mp4',
                        pictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2012/96/iconmonstr-video-8.png',
                        leaf: true
                      }, {
                        id: 'Meeting Audio Record.mp3', label: 'Meeting Audio Record.mp3',
                        pictureUrl: 'http://iconmonstr.com/wp-content/assets/preview/2012/96/iconmonstr-equalizer-1.png',
                        leaf: true
                      }]
                    }
                  ],
                  selectedNodesIDs: this.properties.treeView,
                  allowMultipleSelections: true,
                  allowFoldersSelections: false,
                  nodesPaddingLeft: 15,
                  checkboxEnabled: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'treeViewFieldId'
                }),
                PropertyFieldTagPicker('tags', {
                  label: strings.TagPickerFieldLabel,
                  selectedTags: this.properties.tags,
                  tags: [
                    {key: 'black', name: 'black'},
                    {key: 'blue', name: 'blue'},
                    {key: 'brown', name: 'brown'},
                    {key: 'cyan', name: 'cyan'},
                    {key: 'green', name: 'green'},
                    {key: 'magenta', name: 'magenta'},
                    {key: 'mauve', name: 'mauve'},
                    {key: 'orange', name: 'orange'},
                    {key: 'pink', name: 'pink'},
                    {key: 'purple', name: 'purple'},
                    {key: 'red', name: 'red'},
                    {key: 'rose', name: 'rose'},
                    {key: 'violet', name: 'violet'},
                    {key: 'white', name: 'white'},
                    {key: 'yellow', name: 'yellow'}
                  ],
                  loadingText: 'Loading...',
                  noResultsFoundText: 'No tags found',
                  suggestionsHeaderText: 'Suggested Tags',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'tagsPickerFieldId'
                }),
                PropertyFieldStarRating('starRating', {
                  label: strings.StarRatingFieldLabel,
                  initialValue: this.properties.starRating,
                  starCount: 5,
                  starSize: 24,
                  starColor: '#ffb400',
                  emptyStarColor: '#333',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'starRatingFieldId'
                }),
                PropertyFieldPassword('password', {
                  label: strings.PasswordFieldLabel,
                  initialValue: this.properties.password,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'passwordFieldId'
                }),
                PropertyFieldNumericInput('numeric', {
                  label: strings.NumericInputFieldLabel,
                  initialValue: this.properties.numeric,
                  min: 0,
                  max: 100,
                  step: 1,
                  precision: 0,
                  size: 10,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'passwordFieldId'
                }),
                PropertyFieldRichTextBox('richTextBox', {
                  label: strings.RichTextBoxFieldLabel,
                  initialValue: this.properties.richTextBox,
                  inline: false,
                  minHeight: 100,
                  mode: 'basic', //'basic' or 'standard' or 'full'
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'richFieldId'
                }),
                PropertyFieldDatePicker('date', {
                  label: strings.DateFieldLabel,
                  initialDate: this.properties.date,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'dateFieldId'
                }),
                PropertyFieldDatePicker('date2', {
                  label: strings.DateFieldLabel,
                  initialDate: this.properties.date2,
                  formatDate: this.formatDateIso,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'date2FieldId'
                }),
                PropertyFieldDateTimePicker('datetime', {
                  label: strings.DateTimeFieldLabel,
                  initialDate: this.properties.datetime,
                  //formatDate: this.formatDateIso,
                  timeConvention: ITimeConvention.Hours12,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'dateTimeFieldId'
                }),
                PropertyFieldSliderRange('sliderRange', {
                  label: strings.SliderRangeFieldLabel,
                  initialValue: this.properties.sliderRange,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  showValue: true,
                  disabled: false,
                  min: 0,
                  max: 500,
                  step: 1,
                  orientation: 'horizontal', //'horizontal' or 'vertical'
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'sliderRangeFieldId'
                }),
                PropertyFieldPhoneNumber('phone', {
                  label: strings.PhoneNumberFieldLabel,
                  initialValue: this.properties.phone,
                  phoneNumberFormat: IPhoneNumberFormat.UnitedStates,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'phoneFieldId'
                }),
                PropertyFieldMaskedInput('maskedInput', {
                  label: strings.MaskedInputFieldLabel,
                  initialValue: this.properties.maskedInput,
                  pattern: '\d{4} \d{4} \d{4} \d{4}',
                  placeholder: 'XXXX XXXX XXXX XXXX',
                  maxLength: '19',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'maskedFieldId'
                }),
                PropertyFieldMapPicker('geolocation', {
                  label: strings.GeoLocationFieldLabel,
                  longitude: this.properties.geolocation != null ? this.properties.geolocation.substr(0, this.properties.geolocation.indexOf(",")) : '0',
                  latitude: this.properties.geolocation != null ? this.properties.geolocation.substr(this.properties.geolocation.indexOf(",") + 1, this.properties.geolocation.length - this.properties.geolocation.indexOf(",")) : '0',
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'geoLocationFieldId'
                })
            ],
            },
            {
              groupName: 'SharePoint Fields',
              groupFields: [

                PropertyFieldPicturePicker('picture', {
                  label: strings.PictureFieldLabel,
                  initialValue: this.properties.picture,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  disabled: false,
                  readOnly: true,
                  previewImage: true,
                  allowedFileExtensions: '.gif,.jpg,.jpeg,.bmp,.dib,.tif,.tiff,.ico,.png',
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'pictureFieldId'
                }),
                PropertyFieldDocumentPicker('document', {
                  label: strings.DocumentFieldLabel,
                  initialValue: this.properties.document,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  disabled: false,
                  readOnly: true,
                  previewDocument: true,
                  allowedFileExtensions: '.doc,.docx,.ppt,.pptx,.xls,.xlsx,.pdf,.txt',
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'documentFieldId'
                }),
                PropertyFieldPeoplePicker('people', {
                  label: strings.PeopleFieldLabel,
                  initialData: this.properties.people,
                  allowDuplicate: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'
                }),
                PropertyFieldGroupPicker('groups', {
                  label: strings.GroupFieldLabel,
                  initialData: this.properties.groups,
                  allowDuplicate: false,
                  groupType: IGroupType.SharePoint,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'groupsFieldId'
                }),
               PropertyFieldSPListPicker('list', {
                  label: strings.SPListFieldLabel,
                  selectedList: this.properties.list,
                  includeHidden: false,
                  //baseTemplate: 109,
                  orderBy: PropertyFieldSPListPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listFieldId'
                }),
                PropertyFieldSPFolderPicker('folder', {
                  label: strings.SPFolderFieldLabel,
                  initialFolder: this.properties.folder,
                  //baseFolder: '/sites/devcenter/_catalogs',
                  context: this.context,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'folderFieldId'
                }),
                PropertyFieldSPListMultiplePicker('listsCollection', {
                  label: strings.SPListFieldLabel,
                  selectedLists: this.properties.listsCollection,
                  includeHidden: false,
                  baseTemplate: 109,
                  orderBy: PropertyFieldSPListMultiplePickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listCollectionFieldId'
                })
              ]
            },
            {
              groupName: 'SharePoint Query',
              groupFields: [
                PropertyFieldSPListQuery('query', {
                  label: strings.QueryFieldLabel,
                  query: this.properties.query,
                  includeHidden: false,
                  //baseTemplate: 109,
                  orderBy: PropertyFieldSPListQueryOrderBy.Title,
                  showOrderBy: true,
                  showMax: true,
                  showFilters: true,
                  max: 50,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'spListFieldId'
                }),
                 PropertyFieldDisplayMode('displayMode', {
                  label: strings.DisplayModeFieldLabel,
                  initialValue: this.properties.displayMode,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'displayModeFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}











