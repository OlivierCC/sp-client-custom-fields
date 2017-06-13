/**
 * @file PropertyFieldSortableList.ts
 * Define a custom field of type PropertyFieldSortableList for
 * the SharePoint Framework (SPfx)
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 *
 */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-webpart-base';
import { IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import PropertyFieldSortableListHost, { IPropertyFieldSortableListHostProps } from './PropertyFieldSortableListHost';

/**
 * @enum
 * Sortable List order by enum
 */
export enum ISortableListOrder {
  Key,
  Text
}

/**
 * @interface
 * Public properties of the PropertyFieldSortableList custom field
 *
 */
export interface IPropertyFieldSortableListProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  /**
   * @var
   * Sortable list options
   */
  items: IChoiceGroupOption[];
  /**
   * @var
   * Default selected values of the picker (must be a collection of list keys)
   */
  selectedItems?: string[];
  /**
   * @var
   * Defines the sort by criteria (default is `Text`)
   */
  sortBy?: ISortableListOrder;
  /**
   * @function
   * Defines a onPropertyChange function to raise when the selected value changed.
   * Normally this function must be always defined with the 'this.onPropertyChange'
   * method of the web part object.
   */
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  /**
   * @function
   * This API is called to render the web part.
   * Normally this function must be always defined with the 'this.render.bind(this)'
   * method of the web part object.
   */
  render(): void;
  /**
   * This property is used to indicate the web part's PropertyPane interaction mode: Reactive or NonReactive.
   * The default behaviour is Reactive.
   */
  disableReactivePropertyChanges?: boolean;
  /**
   * @var
   * Parent Web Part properties
   */
  properties: any;
  /**
   * @var
   * An UNIQUE key indicates the identity of this control
   */
  key?: string;
  /**
   * Whether the property pane field is enabled or not.
   */
  disabled?: boolean;
  /**
   * The method is used to get the validation error message and determine whether the input value is valid or not.
   *
   *   When it returns string:
   *   - If valid, it returns empty string.
   *   - If invalid, it returns the error message string and the text field will
   *     show a red border and show an error message below the text field.
   *
   *   When it returns Promise<string>:
   *   - The resolved value is display as error message.
   *   - The rejected, the value is thrown away.
   *
   */
   onGetErrorMessage?: (value: string[]) => string | Promise<string>;
   /**
    * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
    * Default value is 200.
    */
   deferredValidationTime?: number;
}

/**
 * @interface
 * Private properties of the PropertyFieldSortableList custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldSortableList.
 *
 */
export interface IPropertyFieldSortableListPropsInternal extends IPropertyFieldSortableListProps {
  label: string;
  targetProperty: string;
  items: IChoiceGroupOption[];
  selectedItems?: string[];
  sortBy?: ISortableListOrder;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  render(): void;
  disableReactivePropertyChanges?: boolean;
  properties: any;
  key: string;
  disabled?: boolean;
  onGetErrorMessage?: (value: string[]) => string | Promise<string>;
  deferredValidationTime?: number;
}

/**
 * @interface
 * Represents a PropertyFieldSortableList object
 *
 */
class PropertyFieldSortableListBuilder implements IPropertyPaneField<IPropertyFieldSortableListPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldSortableListPropsInternal;

  //Custom properties label: string;
  private label: string;
  private items: IChoiceGroupOption[];
  private selectedItems: string[];
  private sortBy: ISortableListOrder = ISortableListOrder.Text;

  public onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void {}
  private customProperties: any;
  private key: string;
  private disabled: boolean = false;
  private onGetErrorMessage: (value: string[]) => string | Promise<string>;
  private deferredValidationTime: number = 200;
  private renderWebPart: () => void;
  private disableReactivePropertyChanges: boolean = false;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldSortableListPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _targetProperty;
    this.properties = _properties;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.label = _properties.label;
    this.items = _properties.items;
    this.selectedItems = _properties.selectedItems;
    this.onPropertyChange = _properties.onPropertyChange;
    this.customProperties = _properties.properties;
    this.key = _properties.key;
    if (_properties.disabled === true)
      this.disabled = _properties.disabled;
    this.onGetErrorMessage = _properties.onGetErrorMessage;
    if (_properties.deferredValidationTime !== undefined)
      this.deferredValidationTime = _properties.deferredValidationTime;
    if (_properties.sortBy !== undefined)
      this.sortBy = _properties.sortBy;
    this.renderWebPart = _properties.render;
    if (_properties.disableReactivePropertyChanges !== undefined && _properties.disableReactivePropertyChanges != null)
      this.disableReactivePropertyChanges = _properties.disableReactivePropertyChanges;
  }

  /**
   * @function
   * Renders the SPListPicker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldSortableListHostProps> = React.createElement(PropertyFieldSortableListHost, {
      label: this.label,
      targetProperty: this.targetProperty,
      items: this.items,
      selectedItems: this.selectedItems,
      sortBy: this.sortBy,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange,
      properties: this.customProperties,
      key: this.key,
      disabled: this.disabled,
      onGetErrorMessage: this.onGetErrorMessage,
      deferredValidationTime: this.deferredValidationTime,
      render: this.renderWebPart,
      disableReactivePropertyChanges: this.disableReactivePropertyChanges
    });
    //Calls the REACT content generator
    ReactDom.render(element, elem);
  }

  /**
   * @function
   * Disposes the current object
   */
  private dispose(elem: HTMLElement): void {

  }

}

/**
 * @function
 * Helper method to create a SPList Picker on the PropertyPane.
 * @param targetProperty - Target property the SharePoint list picker is associated to.
 * @param properties - Strongly typed SPList Picker properties.
 */
export function PropertyFieldSortableList(targetProperty: string, properties: IPropertyFieldSortableListProps): IPropertyPaneField<IPropertyFieldSortableListPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldSortableListPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      selectedItems: properties.selectedItems,
      items: properties.items,
      sortBy: properties.sortBy,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      onDispose: null,
      onRender: null,
      key: properties.key,
      disabled: properties.disabled,
      onGetErrorMessage: properties.onGetErrorMessage,
      deferredValidationTime: properties.deferredValidationTime,
      render: properties.render,
      disableReactivePropertyChanges: properties.disableReactivePropertyChanges
    };
    //Calls the PropertyFieldSortableList builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldSortableListBuilder(targetProperty, newProperties);
}
