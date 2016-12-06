/**
 * @file PropertyFieldPeoplePicker.ts
 * Define a custom field of type PropertyFieldPeoplePicker for
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
  PropertyPaneFieldType,
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-webpart-base';
import PropertyFieldPeoplePickerHost, { IPropertyFieldPeoplePickerHostProps } from './PropertyFieldPeoplePickerHost';
import { IWebPartContext} from '@microsoft/sp-webpart-base';

/**
 * @interface
 * Defines a People object for the PropertyFieldPeoplePicker
 *
 */
export interface IPropertyFieldPeople {
  /**
   * @var
   * User's full name
   */
  fullName: string;
  /**
   * @var
   * User's login
   */
  login: string;
  /**
   * @var
   * User's email (optional)
   */
  email?: string;
  /**
   * @var
   * User's job title (optional)
   */
  jobTitle?: string;
  /**
   * @var
   * User's initials (optional)
   */
  initials?: string;
  /**
   * @var
   * User's image url (optional)
   */
  imageUrl?: string;
}

/**
 * @interface
 * Public properties of the PropertyFieldPeoplePicker custom field
 *
 */
export interface IPropertyFieldPeoplePickerProps {
  /**
   * @var
   * Property field label
   */
  label: string;
  /**
   * @var
   * Web Part context
   */
  context: IWebPartContext;
  /**
   * @var
   * Intial data to load in the people picker (optional)
   */
  initialData?: IPropertyFieldPeople[];
  /**
   * @var
   * Defines if the People Picker allows to select duplicated users (optional)
   */
  allowDuplicate?: boolean;
  /**
   * @function
   * Defines a onPropertyChange function to raise when the selected value changed.
   * Normally this function must be always defined with the 'this.onPropertyChange'
   * method of the web part object.
   */
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
    /**
   * @var
   * Parent Web Part properties
   */
  properties: any;
}

/**
 * @interface
 * Private properties of the PropertyFieldPeoplePicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldPeoplePicker.
 *
 */
export interface IPropertyFieldPeoplePickerPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  targetProperty: string;
  context: IWebPartContext;
  initialData?: IPropertyFieldPeople[];
  allowDuplicate?: boolean;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
}

/**
 * @interface
 * Represents a PropertyFieldPeoplePicker object
 *
 */
class PropertyFieldPeoplePickerBuilder implements IPropertyPaneField<IPropertyFieldPeoplePickerPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldPeoplePickerPropsInternal;

  //Custom properties
  private label: string;
  private context: IWebPartContext;
  private initialData: IPropertyFieldPeople[];
  private allowDuplicate: boolean = true;
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldPeoplePickerPropsInternal) {
    this.render = this.render.bind(this);
    this.label = _properties.label;
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.onPropertyChange = _properties.onPropertyChange;
    this.context = _properties.context;
    this.initialData = _properties.initialData;
    this.allowDuplicate = _properties.allowDuplicate;
    this.customProperties = _properties.properties;
  }

  /**
   * @function
   * Renders the PeoplePicker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldPeoplePickerHostProps> = React.createElement(PropertyFieldPeoplePickerHost, {
      label: this.label,
      targetProperty: this.targetProperty,
      initialData: this.initialData,
      allowDuplicate: this.allowDuplicate,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange,
      context: this.context,
      properties: this.customProperties
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
 * Helper method to create a People Picker on the PropertyPane.
 * @param targetProperty - Target property the people picker is associated to.
 * @param properties - Strongly typed people Picker properties.
 */
export function PropertyFieldPeoplePicker(targetProperty: string, properties: IPropertyFieldPeoplePickerProps): IPropertyPaneField<IPropertyFieldPeoplePickerPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldPeoplePickerPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      onPropertyChange: properties.onPropertyChange,
      context: properties.context,
      initialData: properties.initialData,
      allowDuplicate: properties.allowDuplicate,
      properties: properties.properties,
      onDispose: null,
      onRender: null
    };
    //Calles the PropertyFieldDatePicker builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldPeoplePickerBuilder(targetProperty, newProperties);
}


