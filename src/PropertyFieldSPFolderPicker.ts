/**
 * @file PropertyFieldSPFolderPicker.ts
 * Define a custom field of type PropertyFieldSPFolderPicker for
 * the SharePoint Framework (SPfx)
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType,
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-webpart-base';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import PropertyFieldSPFolderPickerHost, { IPropertyFieldSPFolderPickerHostProps } from './PropertyFieldSPFolderPickerHost';

/**
 * @interface
 * Public properties of the PropertyFieldSPFolderPicker custom field
 *
 */
export interface IPropertyFieldSPFolderPickerProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  /**
   * @var
   * Defines the initial selected folder
   */
  initialFolder?: string;
  /**
   * @var
   * Defines the root folder. If empty, the base folder is the current web root folder
   */
  baseFolder?: string;
  /**
   * @var
   * Parent web part context
   */
  context: IWebPartContext;
  /**
   * @function
   * Defines a onPropertyChange function to raise when the selected folder changed.
   * Normally this function must be always defined with the 'this.onPropertyChange'
   * method of the web part object.
   */
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
    /**
   * @var
   * Parent Web Part properties
   */
  properties: any;
  /**
   * @var
   * Initial value
   */
  key?: string;
}

/**
 * @interface
 * Private properties of the PropertyFieldSPFolderPicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldSPFolderPicker.
 *
 */
export interface IPropertyFieldSPFolderPickerPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  initialFolder?: string;
  baseFolder?: string;
  targetProperty: string;
  context: IWebPartContext;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
}

/**
 * @interface
 * Represents a PropertyFieldSPFolderPicker object
 *
 */
class PropertyFieldSPFolderPickerBuilder implements IPropertyPaneField<IPropertyFieldSPFolderPickerPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldSPFolderPickerPropsInternal;

  //Custom properties
  private label: string;
  private initialFolder: string;
  private baseFolder: string;
  private context: IWebPartContext;
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;
  private key: string;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldSPFolderPickerPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    this.initialFolder = _properties.initialFolder;
    this.baseFolder = this.baseFolder;
    this.context = _properties.context;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.onPropertyChange = _properties.onPropertyChange;
    this.customProperties = _properties.properties;
    this.key = _properties.key;
  }

  /**
   * @function
   * Renders the SPFolderPicker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldSPFolderPickerHostProps> = React.createElement(PropertyFieldSPFolderPickerHost, {
      label: this.label,
      initialFolder: this.initialFolder,
      baseFolder: this.baseFolder,
      context: this.context,
      targetProperty: this.targetProperty,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange,
      properties: this.customProperties,
      key: this.key
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
 * Helper method to create a SharePoint Folder Picker on the PropertyPane.
 * @param targetProperty - Target property the Folder picker is associated to.
 * @param properties - Strongly typed Folder Picker properties.
 */
export function PropertyFieldSPFolderPicker(targetProperty: string, properties: IPropertyFieldSPFolderPickerProps): IPropertyPaneField<IPropertyFieldSPFolderPickerPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldSPFolderPickerPropsInternal = {
      label: properties.label,
      initialFolder: properties.initialFolder,
      baseFolder: properties.baseFolder,
      context: properties.context,
      targetProperty: targetProperty,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      onDispose: null,
      onRender: null,
      key: properties.key
    };
    //Calles the PropertyFieldSPFolderPicker builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldSPFolderPickerBuilder(targetProperty, newProperties);
}


