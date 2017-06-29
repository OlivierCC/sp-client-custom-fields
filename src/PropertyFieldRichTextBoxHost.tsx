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
  keyCopy: string;
}


export interface IPropertyFieldRichTextBoxHostState {
}

/**
 * @class
 * Renders the controls for PropertyFieldRichTextBox component
 */
export default class PropertyFieldRichTextBoxHost extends React.Component<IPropertyFieldRichTextBoxHostProps, IPropertyFieldRichTextBoxHostState> {

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldRichTextBoxHostProps) {
    super(props);

    //Bind the current object to the external called onSelectDate method
  }


  /**
   * @function
   * Renders the controls
   */
  public render(): JSX.Element {
    //Renders content
    var minHeight = 100;
    if (this.props.minHeight != null)
      minHeight = this.props.minHeight;
    return (
      <div>
        <Label>{this.props.label}</Label>
        <div style={{border: '1px solid #c8c8c8', minHeight: minHeight + 'px'}}>
          <textarea disabled={this.props.disabled}
            name={this.props.keyCopy + '-' + this.props.context.instanceId + '-editor'}
            id={this.props.keyCopy + '-' + this.props.context.instanceId + '-editor'}
            defaultValue={this.props.initialValue}
            ></textarea>
        </div>
        <div>
            <div aria-live='assertive' className='ms-u-screenReaderOnly' data-automation-id='error-message'>
              <span id={this.props.keyCopy + '-' + this.props.context.instanceId + '-errorMssg1'}/>
            </div>
            <span>
              <p className='ms-TextField-errorMessage ms-u-slideDownIn20'>
                <span id={this.props.keyCopy + '-' + this.props.context.instanceId + '-errorMssg2'}/>
              </p>
            </span>
        </div>
      </div>
    );
  }
}
