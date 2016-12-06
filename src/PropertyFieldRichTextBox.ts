/**
 * @file PropertyFieldRichTextBox.ts
 * Define a custom field of type PropertyFieldRichTextBox for
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
import PropertyFieldRichTextBoxHost, { IPropertyFieldRichTextBoxHostProps } from './PropertyFieldRichTextBoxHost';
import ModuleLoader from '@microsoft/sp-module-loader';

/**
 * @interface
 * Public properties of the PropertyFieldRichTextBox custom field
 *
 */
export interface IPropertyFieldRichTextBoxProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  /**
   * @var
   * Initial value
   */
  initialValue?: string;
  /**
   * @var
   * 'basic' or 'standard' or 'full'. Default is basic
   */
  mode?: string;
  /**
   * @var
   * Popin toolbar or Classic toolbar
   */
  inline?: boolean;
  /**
   * @var
   * Textarea min height
   */
  minHeight?: number;
  /**
   * @function
   * Defines a onPropertyChange function to raise when the selected Color changed.
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
 * Private properties of the PropertyFieldRichTextBox custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldRichTextBox.
 *
 */
export interface IPropertyFieldRichTextBoxPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  initialValue?: string;
  targetProperty: string;
  mode?: string;
  inline?: boolean;
  minHeight?: number;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
}

/**
 * @interface
 * Represents a PropertyFieldRichTextBox object
 *
 */
class PropertyFieldRichTextBoxBuilder implements IPropertyPaneField<IPropertyFieldRichTextBoxPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldRichTextBoxPropsInternal;

  //Custom properties
  private label: string;
  private initialValue: string;
  private mode: string;
  private  inline: boolean;
  private minHeight: number;
  private guid: string;
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldRichTextBoxPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    this.mode = _properties.mode;
    this.inline = _properties.inline;
    this.initialValue = _properties.initialValue;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.minHeight = this.minHeight;
    this.onPropertyChange = _properties.onPropertyChange;
    this.render = this.render.bind(this);
    this.customProperties = _properties.properties;
    this.guid = this.getGuid();
  }

  private getGuid(): string {
    return this.s4() + this.s4() + '-' + this.s4() + '-' + this.s4() + '-' +
      this.s4() + '-' + this.s4() + this.s4() + this.s4();
  }

  private s4(): string {
      return Math.floor((1 + Math.random()) * 0x10000)
        .toString(16)
        .substring(1);
    }

  /**
   * @function
   * Renders the ColorPicker field content
   */
  private render(elem: HTMLElement): void {

    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldRichTextBoxHostProps> = React.createElement(PropertyFieldRichTextBoxHost, {
      label: this.label,
      initialValue: this.initialValue,
      targetProperty: this.targetProperty,
      mode: this.mode,
      inline: this.inline,
      minHeight: this.minHeight,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange,
      guid: this.guid,
      properties: this.customProperties
    });
    //Calls the REACT content generator
    ReactDom.render(element, elem);

    var fMode = 'basic';
    if (this.mode != null)
      fMode = this.mode;
    var ckEditorCdn = '//cdn.ckeditor.com/4.5.11/{0}/ckeditor.js'.replace("{0}", fMode);
    ModuleLoader.loadScript(ckEditorCdn, 'CKEDITOR').then((CKEDITOR: any): void => {
      if (this.inline == null || this.inline === false)
        CKEDITOR.replace( this.guid + '-editor', {
            skin: 'kama,//cdn.ckeditor.com/4.4.3/full-all/skins/kama/'
        }  );
      else
        CKEDITOR.inline( this.guid + '-editor', {
            skin: 'kama,//cdn.ckeditor.com/4.4.3/full-all/skins/kama/'
        }   );

      for (var i in CKEDITOR.instances) {
        CKEDITOR.instances[i].on('change', (elm?, val?) =>
        {
          CKEDITOR.instances[i].updateElement();
          var value = ((document.getElementById(this.guid + '-editor')) as any).value;
          if (this.onPropertyChange && value != null) {
            this.customProperties[this.targetProperty] = value;
            this.onPropertyChange(this.targetProperty, this.properties.initialValue, value);
          }
        });
      }
    });
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
 * Helper method to create a Color Picker on the PropertyPane.
 * @param targetProperty - Target property the Color picker is associated to.
 * @param properties - Strongly typed Color Picker properties.
 */
export function PropertyFieldRichTextBox(targetProperty: string, properties: IPropertyFieldRichTextBoxProps): IPropertyPaneField<IPropertyFieldRichTextBoxPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldRichTextBoxPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      initialValue: properties.initialValue,
      mode: properties.mode,
      inline: properties.inline,
      minHeight: properties.minHeight,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      onDispose: null,
      onRender: null
    };
    //Calles the PropertyFieldRichTextBox builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldRichTextBoxBuilder(targetProperty, newProperties);
}


