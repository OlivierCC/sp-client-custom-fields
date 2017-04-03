/**
 * @file PropertyFieldAutoCompleteHost.tsx
 * Renders the controls for PropertyFieldAutoComplete component
 *
 * @copyright 2017 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldAutoCompletePropsInternal } from './PropertyFieldAutoComplete';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Async } from 'office-ui-fabric-react/lib/Utilities';

/**
 * @interface
 * PropertyFieldAutoCompleteHost properties interface
 *
 */
export interface IPropertyFieldAutoCompleteHostProps extends IPropertyFieldAutoCompletePropsInternal {
}

export interface IPropertyFieldAutoCompleteState {
  currentValue?: string;
  suggestions: string[];
  isOpen: boolean;
  hover: string;
  keyPosition: number;
  isHoverDropdown: boolean;
  errorMessage: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldAutoComplete component
 */
export default class PropertyFieldAutoCompleteHost extends React.Component<IPropertyFieldAutoCompleteHostProps, IPropertyFieldAutoCompleteState> {

  private async: Async;
  private delayedValidate: (value: string) => void;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldAutoCompleteHostProps) {
    super(props);

    this.async = new Async(this);
    this.state = {
       keyPosition: -1,
       errorMessage: '',
       isOpen: false,
       isHoverDropdown: false,
       hover: '',
       currentValue: this.props.initialValue !== undefined ? this.props.initialValue : '',
       suggestions: this.props.suggestions
    };

    //Bind the current object to the external called onSelectDate method
    this.onValueChanged = this.onValueChanged.bind(this);
    this.onOpenDialog = this.onOpenDialog.bind(this);
    this.toggleHover = this.toggleHover.bind(this);
    this.getSuggestions = this.getSuggestions.bind(this);
    this.toggleHoverLeave = this.toggleHoverLeave.bind(this);
    this.onClickItem = this.onClickItem.bind(this);
    this.onInputKeyDown = this.onInputKeyDown.bind(this);
    this.onInputBlur = this.onInputBlur.bind(this);
    this.onInputKeyPress = this.onInputKeyPress.bind(this);
    this.onInputKeyUp = this.onInputKeyUp.bind(this);
    this.onClickInput = this.onClickInput.bind(this);
    this.mouseEnterDropDown = this.mouseEnterDropDown.bind(this);
    this.mouseLeaveDropDown = this.mouseLeaveDropDown.bind(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
  }

  /**
   * @function
   * Function called when the component value changed
   */
  private onValueChanged(element: any): void {
    //Checks if there is a method to called
    if (this.props.onPropertyChange && element != null) {
      var newValue: string = element.currentTarget.value;
      this.state.currentValue = newValue;
      this.state.keyPosition = -1;
      this.state.isOpen = true;
      this.state.suggestions = this.getSuggestions(newValue);
      this.setState(this.state);
      this.delayedValidate(newValue);
    }
  }

  private getSuggestions(value: string) {
    if (value == '') {
      return this.props.suggestions;
    }
    const escapeRegexCharacters = str => str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const escapedValue = escapeRegexCharacters(value.trim());
    if (escapedValue === '') {
      return [];
    }
    const regex = new RegExp('^' + escapedValue, 'i');
    return this.props.suggestions.filter(language => regex.test(language));
  };

  private onInputBlur(elm?: any) {
    if (this.state.hover == '') {
      this.state.isOpen = false;
      this.state.hover = '';
      this.state.keyPosition = -1;
      this.setState(this.state);
    }
  }

  private onInputKeyPress(elm?: any) {
    if (elm.keyCode != 40 && elm.keyCode != 38) {
      this.state.keyPosition = -1;
      this.state.hover = '';
      this.setState(this.state);
    }
    if (elm.keyCode == 0) {
      this.state.isOpen = false;
      this.state.hover = '';
      this.state.keyPosition = -1;
      this.setState(this.state);
    }
  }

  private onInputKeyDown(elm?: any) {
    if (elm.keyCode === 40) {
      this.state.keyPosition = this.state.keyPosition + 1;
      if (this.state.keyPosition >= this.state.suggestions.length)
        this.state.keyPosition = this.state.suggestions.length - 1;
      this.state.currentValue = this.state.suggestions[this.state.keyPosition];
      this.setState(this.state);
      this.delayedValidate(this.state.currentValue);
    }
  }

  private onInputKeyUp(elm?: any) {
    if (elm.keyCode === 38) {
      this.state.keyPosition = this.state.keyPosition - 1;
      if (this.state.keyPosition < 0)
        this.state.keyPosition = 0;
      this.state.currentValue = this.state.suggestions[this.state.keyPosition];
      this.setState(this.state);
      this.delayedValidate(this.state.currentValue);
    }
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: string): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.initialValue, value);
      return;
    }

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || '');
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.initialValue, value);
        this.setState({ errorMessage: result} as IPropertyFieldAutoCompleteState);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.initialValue, value);
          this.setState({ errorMessage } as IPropertyFieldAutoCompleteState);
        });
      }
    }
    else {
      this.notifyAfterValidate(this.props.initialValue, value);
    }
  }

  /**
   * @function
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: string, newValue: string) {
    this.props.properties[this.props.targetProperty] = newValue;
    this.props.onPropertyChange(this.props.targetProperty, oldValue, newValue);
  }

  /**
   * @function
   * Called when the component will unmount
   */
  public componentWillUnmount() {
    if (this.async !== undefined)
      this.async.dispose();
  }

    /**
   * @function
   * Function to open the dialog
   */
  private onOpenDialog(): void {
    if (this.props.disabled === true)
      return;
    this.state.isOpen = !this.state.isOpen;
    this.setState(this.state);
  }

  /**
   * @function
   * Mouse is hover a font
   */
  private toggleHover(element?: any) {
    var hoverFont: string = element.currentTarget.textContent;
    this.state.hover = hoverFont;
    this.setState(this.state);
  }

  /**
   * @function
   * Mouse is leaving a font
   */
  private toggleHoverLeave(element?: any) {
    this.state.hover = '';
    this.setState(this.state);
  }

  /**
   * @function
   * Mouse is hover the fontpicker
   */
  private mouseEnterDropDown(element?: any) {
    this.state.isHoverDropdown = true;
    this.setState(this.state);
  }

  /**
   * @function
   * Mouse is leaving the fontpicker
   */
  private mouseLeaveDropDown(element?: any) {
    this.state.isHoverDropdown = false;
    this.setState(this.state);
  }

  /**
   * @function
   * User clicked on a font
   */
  private onClickItem(element?: any) {
    element.stopPropagation();
    var clickedFont: string = element.currentTarget.textContent;
    this.state.currentValue = clickedFont;
    this.onOpenDialog();
    this.delayedValidate(clickedFont);
  }

  private onClickInput(elm?: any) {
    this.state.isOpen = true;
    this.state.suggestions = this.getSuggestions(this.state.currentValue);
    this.setState(this.state);
  }


  /**
   * @function
   * Renders the controls
   */
  public render(): JSX.Element {

    var fontSelect = {
        fontSize: '16px',
        width: '100%',
        position: 'relative',
        display: 'inline-block',
        zoom: 1
      };
      var dropdownColor = '1px solid #c8c8c8';
      if (this.props.disabled === true)
        dropdownColor = '1px solid #f4f4f4';
      else if (this.state.isOpen === true)
        dropdownColor = '1px solid #3091DE';
      else if (this.state.isHoverDropdown === true)
        dropdownColor = '1px solid #767676';
      var fontSelectA = {
        backgroundColor: this.props.disabled === true ? '#f4f4f4' : '#fff',
        borderRadius        : '0px',
        backgroundClip        : 'padding-box',
        border: dropdownColor,
        display: 'block',
        overflow: 'hidden',
        whiteSpace: 'nowrap',
        position: 'relative',
        height: '26px',
        lineHeight: '26px',
        padding: '0 0 0 8px',
        color: this.props.disabled === true ? '#a6a6a6' : '#444',
        textDecoration: 'none',
        cursor: this.props.disabled === true ? 'default' : 'pointer'
      };
      var fontSelectASpan = {
        marginRight: '26px',
        display: 'block',
        overflow: 'hidden',
        whiteSpace: 'nowrap',
        lineHeight: '1.8',
        textOverflow: 'ellipsis',
        cursor: this.props.disabled === true ? 'default' : 'pointer',
        //fontFamily: this.state.safeSelectedFont != null && this.state.safeSelectedFont != '' ? this.state.safeSelectedFont : 'Arial',
        //fontSize: this.state.safeSelectedFont,
        fontWeight: 400
      };
      var fontSelectADiv = {
        borderRadius        : '0 0px 0px 0',
        backgroundClip        : 'padding-box',
        border: '0px',
        position: 'absolute',
        right: '0',
        top: '0',
        display: 'block',
        height: '100%',
        width: '22px'
      };
      var fontSelectADivB = {
        display: 'block',
        width: '100%',
        height: '100%',
        cursor: this.props.disabled === true ? 'default' : 'pointer',
        marginTop: '2px'
      };
      var fsDrop = {
        background: '#fff',
        border: '1px solid #aaa',
        borderTop: '0',
        position: 'absolute',
        top: '32px',
        left: '0',
        width: 'calc(100% - 2px)',
        //boxShadow: '0 4px 5px rgba(0,0,0,.15)',
        zIndex: 999,
        display: this.props.disabled === true ? 'none' :  this.state.isOpen ? 'block' : 'none'
      };
      var fsResults = {
        margin: '0 4px 4px 0',
        maxHeight: '190px',
        width: 'calc(100% - 4px)',
        padding: '0 0 0 4px',
        position: 'relative',
        overflowX: 'hidden',
        overflowY: 'auto'
      };
      var carret: string = this.state.isOpen ? 'ms-Icon ms-Icon--ChevronUp' : 'ms-Icon ms-Icon--ChevronDown';
      //Renders content
      return (
        <div style={{ marginBottom: '8px'}}>
          <Label>{this.props.label}</Label>
          <div style={fontSelect}>
            <input disabled={this.props.disabled}
              placeholder={this.props.placeHolder !== undefined ? this.props.placeHolder : ''}
              label={this.props.label}
              value={this.state.currentValue}
              className="ms-TextField-field"
              onClick={this.onClickInput}
              onBlur={this.onInputBlur}
              onKeyUp={this.onInputKeyUp}
              onKeyDown={this.onInputKeyDown}
              onKeyPress={this.onInputKeyPress}
              onChange={this.onValueChanged}
              aria-invalid={ !!this.state.errorMessage }
              />
            <div style={fsDrop}>
              <ul style={fsResults}>
                {this.state.suggestions.map((sug: string, index: number) => {
                  var backgroundColor: string = 'transparent';
                  if (this.state.currentValue === sug)
                    backgroundColor = '#c7e0f4';
                  else if (this.state.hover === sug)
                    backgroundColor = '#eaeaea';
                  var innerStyle = {
                    lineHeight: '80%',
                    padding: '7px 7px 8px',
                    margin: '0',
                    listStyle: 'none',
                    backgroundColor: backgroundColor,
                    cursor: 'pointer'
                  };
                  return (
                    <li key={'autocompletepicker-' + index} role="menuitem" onMouseEnter={this.toggleHover} onClick={this.onClickItem} onMouseLeave={this.toggleHoverLeave} style={innerStyle}>{sug}</li>
                  );
                })
                }
              </ul>
            </div>
          </div>
          { this.state.errorMessage != null && this.state.errorMessage != '' && this.state.errorMessage != undefined ?
              <div><div aria-live='assertive' className='ms-u-screenReaderOnly' data-automation-id='error-message'>{  this.state.errorMessage }</div>
              <span>
                <p className='ms-TextField-errorMessage ms-u-slideDownIn20'>{ this.state.errorMessage }</p>
              </span>
              </div>
            : ''}
        </div>
      );
  }
}
