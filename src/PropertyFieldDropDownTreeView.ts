/**
 * @file PropertyFieldDropDownTreeView.ts
 * Define a custom field of type PropertyFieldDropDownTreeView for
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
import PropertyFieldDropDownTreeViewHost, { IPropertyFieldDropDownTreeViewHostProps } from './PropertyFieldDropDownTreeViewHost';

/**
 * @interface
 * Describes a treeview Node
 */
export interface IDropDownTreeViewNode {
  /**
   * @var
   * Node's ID (must be unique)
   */
  id: string;
  /**
   * @var
   * Node's label
   */
  label: string;
  /**
   * @var
   * Defines if the node is collapsed or not
   */
  collapsed?: boolean;
  /**
   * @var
   * Defines if the node is a leaf
   */
  leaf?: boolean;
  /**
   * @var
   * Array of child nodes
   */
  children?: IDropDownTreeViewNode[];
  /**
   * @var
   * Node's picture URL. Note: the image size will be always 18px x 18px.
   */
  pictureUrl?: string;
  /**
   * @var
   * Node's picture URL used when the node is selected. If empty, the pictureUrl will be used. Note: the image size will be always 18px x 18px.
   */
  selectedPictureUrl?: string;
  /**
   * @var
   * Node's picutre URL used when the node is a folder and when the node is expanded. If empty, the pictureUrl will be used. Note: the image size will be always 18px x 18px.
   */
  expandedPictureUrl?: string;
}

/**
 * @interface
 * Public properties of the PropertyFieldDropDownTreeView custom field
 *
 */
export interface IPropertyFieldDropDownTreeViewProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  /**
   * @var
   * Tree root nodes
   */
  tree: IDropDownTreeViewNode[];
  /**
   * @var
   * Selected nodes ids
   */
  selectedNodesIDs?: string[];
  /**
   * @var
   * Defines if the user can select multiple nodes or not (default is `false')
   */
  allowMultipleSelections?: boolean;
  /**
   * @var
   * Defines if the user can select a folder or only leaf (default is `true`)
   */
  allowFoldersSelections?: boolean;
  /**
   * @var
   * Defines the nodes padding left. Default value is `20` pixels
   */
  nodesPaddingLeft?: number;
  /**
   * @var
   * Defines if a checkbox is displayed with the item (default is `true`)
   */
  checkboxEnabled?: boolean;
  /**
   * @function
   * Defines a onPropertyChange function to raise when the selected Font changed.
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
 * Private properties of the PropertyFieldDropDownTreeView custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldDropDownTreeView.
 *
 */
export interface IPropertyFieldDropDownTreeViewPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  targetProperty: string;
  tree: IDropDownTreeViewNode[];
  selectedNodesIDs?: string[];
  allowMultipleSelections?: boolean;
  allowFoldersSelections?: boolean;
  nodesPaddingLeft?: number;
  checkboxEnabled?: boolean;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  render(): void;
  disableReactivePropertyChanges?: boolean;
  properties: any;
  disabled?: boolean;
  onGetErrorMessage?: (value: string[]) => string | Promise<string>;
  deferredValidationTime?: number;
}

/**
 * @interface
 * Represents a PropertyFieldDropDownTreeView object
 *
 */
class PropertyFieldDropDownTreeViewBuilder implements IPropertyPaneField<IPropertyFieldDropDownTreeViewPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldDropDownTreeViewPropsInternal;

  //Custom properties
  private label: string;
  private tree: IDropDownTreeViewNode[];
  private selectedNodesIDs: string[] = [];
  private allowMultipleSelections: boolean = false;
  private allowFoldersSelections: boolean = true;
  private nodesPaddingLeft: number = 20;
  private checkboxEnabled: boolean = true;
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
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
  public constructor(_targetProperty: string, _properties: IPropertyFieldDropDownTreeViewPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    this.tree = _properties.tree;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.onPropertyChange = _properties.onPropertyChange;
    this.customProperties = _properties.properties;
    this.key = _properties.key;
    if (_properties.disabled === true)
      this.disabled = _properties.disabled;
    this.onGetErrorMessage = _properties.onGetErrorMessage;
    if (_properties.deferredValidationTime !== undefined)
      this.deferredValidationTime = _properties.deferredValidationTime;
    if (_properties.selectedNodesIDs !== undefined && _properties.selectedNodesIDs != null)
      this.selectedNodesIDs = _properties.selectedNodesIDs;
    if (_properties.allowMultipleSelections !== undefined)
      this.allowMultipleSelections = _properties.allowMultipleSelections;
    if (_properties.allowFoldersSelections !== undefined)
      this.allowFoldersSelections = _properties.allowFoldersSelections;
    if (_properties.nodesPaddingLeft !== undefined && _properties.nodesPaddingLeft != null)
      this.nodesPaddingLeft = _properties.nodesPaddingLeft;
    if (_properties.checkboxEnabled !== undefined && _properties.checkboxEnabled != null)
      this.checkboxEnabled = _properties.checkboxEnabled;
    this.renderWebPart = _properties.render;
    if (_properties.disableReactivePropertyChanges !== undefined && _properties.disableReactivePropertyChanges != null)
      this.disableReactivePropertyChanges = _properties.disableReactivePropertyChanges;
}

  /**
   * @function
   * Renders the field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldDropDownTreeViewHostProps> = React.createElement(PropertyFieldDropDownTreeViewHost, {
      label: this.label,
      tree: this.tree,
      selectedNodesIDs: this.selectedNodesIDs,
      allowMultipleSelections: this.allowMultipleSelections,
      allowFoldersSelections: this.allowFoldersSelections,
      nodesPaddingLeft: this.nodesPaddingLeft,
      checkboxEnabled: this.checkboxEnabled,
      targetProperty: this.targetProperty,
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
 * Helper method to create a Font Picker on the PropertyPane.
 * @param targetProperty - Target property the Font picker is associated to.
 * @param properties - Strongly typed Font Picker properties.
 */
export function PropertyFieldDropDownTreeView(targetProperty: string, properties: IPropertyFieldDropDownTreeViewProps): IPropertyPaneField<IPropertyFieldDropDownTreeViewPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldDropDownTreeViewPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      tree: properties.tree,
      selectedNodesIDs: properties.selectedNodesIDs,
      allowMultipleSelections: properties.allowMultipleSelections,
      allowFoldersSelections: properties.allowFoldersSelections,
      nodesPaddingLeft: properties.nodesPaddingLeft,
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
    //Calls the PropertyFieldDropDownTreeView builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldDropDownTreeViewBuilder(targetProperty, newProperties);
}


