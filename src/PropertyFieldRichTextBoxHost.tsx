/**
 * @file PropertyFieldRichTextBoxHost.tsx
 * Renders the controls for PropertyFieldRichTextBox component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldRichTextBoxPropsInternal } from './PropertyFieldRichTextBox';
import { Label } from 'office-ui-fabric-react/lib/Label';

/**
 * @interface
 * PropertyFieldRichTextBoxHost properties interface
 *
 */
export interface IPropertyFieldRichTextBoxHostProps extends IPropertyFieldRichTextBoxPropsInternal {
}


export interface IPropertyFieldRichTextBoxHostState {
  scripLoaded: boolean;
}

/**
 * @class
 * Renders the controls for PropertyFieldRichTextBox component
 */
export default class PropertyFieldRichTextBoxHost extends React.Component<IPropertyFieldRichTextBoxHostProps, IPropertyFieldRichTextBoxHostState> {

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldRichTextBoxHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method

    //this.setState({scripLoaded: false});
  }


  /**
   * @function
   * Renders the datepicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {
    //Renders content
    var minHeight = 100;
    if (this.props.minHeight != null)
      minHeight = this.props.minHeight;
    return (
      <div>
        <Label>{this.props.label}</Label>
        <div style={{border: '1px solid #c8c8c8', minHeight: minHeight + 'px'}}><textarea name="editor1" id="editor1">{this.props.initialValue}</textarea></div>
      </div>
    );
  }
}
