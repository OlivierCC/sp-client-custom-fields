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
import { SPComponentLoader } from '@microsoft/sp-loader';
import { Async } from 'office-ui-fabric-react/lib/Utilities';

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
   onGetErrorMessage?: (value: string) => string | Promise<string>;
   /**
    * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
    * Default value is 200.
    */
   deferredValidationTime?: number;
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
  disabled?: boolean;
  onGetErrorMessage?: (value: string) => string | Promise<string>;
  deferredValidationTime?: number;
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
  private inline: boolean;
  private minHeight: number;
  private guid: string;
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;
  private key: string;
  private disabled: boolean = false;
  private onGetErrorMessage: (value: string) => string | Promise<string>;
  private deferredValidationTime: number = 200;

  private latestValidateValue: string;
  private async: Async;
  private delayedValidate: (value: string) => void;

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
    this.key = _properties.key;
    if (_properties.disabled === true)
      this.disabled = _properties.disabled;
    this.onGetErrorMessage = _properties.onGetErrorMessage;
    if (_properties.deferredValidationTime !== undefined)
      this.deferredValidationTime = _properties.deferredValidationTime;

    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.deferredValidationTime);
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
      properties: this.customProperties,
      key: this.key,
      disabled: this.disabled,
      onGetErrorMessage: this.onGetErrorMessage,
      deferredValidationTime: this.deferredValidationTime
    });
    //Calls the REACT content generator
    ReactDom.render(element, elem);

    var fMode = 'basic';
    if (this.mode != null)
      fMode = this.mode;
    var ckEditorCdn = '//cdn.ckeditor.com/4.5.11/{0}/ckeditor.js'.replace("{0}", fMode);

    SPComponentLoader.loadScript(ckEditorCdn, { globalExportsName: 'CKEDITOR' }).then((CKEDITOR: any): void => {
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
          this.delayedValidate(value);
        });
      }
    });
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: string): void {
    if (this.onGetErrorMessage === null || this.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.initialValue, value);
      return;
    }

    if (this.latestValidateValue === value)
      return;
    this.latestValidateValue = value;

    var result: string | PromiseLike<string> = this.onGetErrorMessage(value || '');
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.initialValue, value);
        ((document.getElementById(this.guid + '-errorMssg1')) as any).innerHTML = result;
        ((document.getElementById(this.guid + '-errorMssg2')) as any).innerHTML = result;
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.initialValue, value);
          ((document.getElementById(this.guid + '-errorMssg1')) as any).innerHTML = errorMessage;
          ((document.getElementById(this.guid + '-errorMssg2')) as any).innerHTML = errorMessage;
        });
      }
    }
    else {
      this.notifyAfterValidate(this.initialValue, value);
    }
  }

  /**
   * @function
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: string, newValue: string) {
    if (this.onPropertyChange && newValue != null) {
      this.customProperties[this.targetProperty] = newValue;
      this.onPropertyChange(this.targetProperty, this.properties.initialValue, newValue);
    }
  }

  /**
   * @function
   * Disposes the current object
   */
  private dispose(elem: HTMLElement): void {
    this.async.dispose();
  }

}

/**
 * @function
 * Helper method to create the customer field on the PropertyPane.
 * @param targetProperty - Target property the custom field is associated to.
 * @param properties - Strongly typed custom field properties.
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
      onRender: null,
      key: properties.key,
      disabled: properties.disabled,
      onGetErrorMessage: properties.onGetErrorMessage,
      deferredValidationTime: properties.deferredValidationTime
    };
    //Calls the PropertyFieldRichTextBox builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldRichTextBoxBuilder(targetProperty, newProperties);
}


