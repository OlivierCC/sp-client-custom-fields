/**
 * @file PropertyFieldFontPickerHost.tsx
 * Renders the controls for PropertyFieldFontPicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldFontPickerPropsInternal } from './PropertyFieldFontPicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
import GuidHelper from './GuidHelper';

/**
 * @interface
 * PropertyFieldFontPickerHost properties interface
 *
 */
export interface IPropertyFieldFontPickerHostProps extends IPropertyFieldFontPickerPropsInternal {
}

/**
 * @interface
 * PropertyFieldFontPickerHost state interface
 *
 */
export interface IPropertyFieldFontPickerHostState {
  isOpen: boolean;
  isHoverDropdown?: boolean;
  hoverFont?: string;
  selectedFont?: string;
  safeSelectedFont?: string;
  errorMessage?: string;
}

/**
 * @interface
 * Define a safe font object
 *
 */
interface ISafeFont {
  Name: string;
  SafeValue: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldFontPicker component
 */
export default class PropertyFieldFontPickerHost extends React.Component<IPropertyFieldFontPickerHostProps, IPropertyFieldFontPickerHostState> {

  /**
   * @var
   * Defines the font series
   */
  private fonts: ISafeFont[] = [
    {Name: "Andale Mono", SafeValue: '"Andale Mono",AndaleMono,monospace'},
    {Name: "Arial", SafeValue: 'Arial,""Helvetica Neue",Helvetica,sans-serif'},
    {Name: "Arial Black", SafeValue: '"Arial Black","Arial Bold",Gadget,sans-serif'},
    {Name: "Arial Narrow", SafeValue: '"Arial Narrow",Arial,sans-serif'},
    {Name: "Arial Rounded MT Bold", SafeValue: '"Arial Rounded MT Bold","Helvetica Rounded",Arial,sans-serif'},
    {Name: "Avant Garde", SafeValue: '"Avant Garde",Avantgarde,"Century Gothic",CenturyGothic,AppleGothic,sans-serif'},
    {Name: "Baskerville", SafeValue: 'Baskerville,"Baskerville Old Face","Hoefler Text",Garamond,"Times New Roman",serif'},
    {Name: "Big Caslon", SafeValue: '"Big Caslon","Book Antiqua","Palatino Linotype",Georgia,serif'},
    {Name: "Bodoni MT", SafeValue: '"Bodoni MT",Didot,"Didot LT STD","Hoefler Text",Garamond,"Times New Roman",serif'},
    {Name: "Book Antiqua", SafeValue: '"Book Antiqua",Palatino,"Palatino Linotype","Palatino LT STD",Georgia,serif'},
    {Name: "Brush Script MT", SafeValue: '"Brush Script MT",cursive'},
    {Name: "Calibri", SafeValue: 'Calibri,Candara,Segoe,"Segoe UI",Optima,Arial,sans-serif'},
    {Name: "Calisto MT", SafeValue: '"Calisto MT","Bookman Old Style",Bookman,"Goudy Old Style",Garamond,"Hoefler Text","Bitstream Charter",Georgia,serif'},
    {Name: "Cambria", SafeValue: 'Cambria,Georgia,serif'},
    {Name: "Candara", SafeValue: 'Candara,Calibri,Segoe,"Segoe UI",Optima,Arial,sans-serif'},
    {Name: "Century Gothic", SafeValue: '"Century Gothic",CenturyGothic,AppleGothic,sans-serif'},
    {Name: "Consolas", SafeValue: 'Consolas,monaco,monospace'},
    {Name: "Copperplate", SafeValue: 'Copperplate,"Copperplate Gothic Light",fantasy'},
    {Name: "Courier New", SafeValue: '"Courier New",Courier,"Lucida Sans Typewriter","Lucida Typewriter",monospace'},
    {Name: "Didot", SafeValue: 'Didot,"Didot LT STD","Hoefler Text",Garamond,"Times New Roman",serif'},
    {Name: "Franklin Gothic Medium", SafeValue: '"Franklin Gothic Medium","Franklin Gothic","ITC Franklin Gothic",Arial,sans-serif'},
    {Name: "Futura", SafeValue: 'Futura,"Trebuchet MS",Arial,sans-serif'},
    {Name: "Garamond", SafeValue: 'Garamond,Baskerville,"Baskerville Old Face","Hoefler Text","Times New Roman",serif'},
    {Name: "Geneva", SafeValue: 'Geneva,Tahoma,Verdana,sans-serif'},
    {Name: "Georgia", SafeValue: 'Georgia,Times,"Times New Roman",serif'},
    {Name: "Gill Sans", SafeValue: '"Gill Sans","Gill Sans MT",Calibri,sans-serif'},
    {Name: "Goudy Old Style", SafeValue: '"Goudy Old Style",Garamond,"Big Caslon","Times New Roman",serif'},
    {Name: "Helvetica", SafeValue: '"Helvetica Neue",Helvetica,Arial,sans-serif'},
    {Name: "Hoefler Text", SafeValue: '"Hoefler Text","Baskerville Old Face",Garamond,"Times New Roman",serif'},
    {Name: "Impact", SafeValue: 'Impact,Haettenschweiler,"Franklin Gothic Bold",Charcoal,"Helvetica Inserat","Bitstream Vera Sans Bold","Arial Black","sans serif"'},
    {Name: "Lucida Bright", SafeValue: '"Lucida Bright",Georgia,serif'},
    {Name: "Lucida Console", SafeValue: '"Lucida Console","Lucida Sans Typewriter",monaco,"Bitstream Vera Sans Mono",monospace'},
    {Name: "Lucida Grande", SafeValue: '"Lucida Grande","Lucida Sans Unicode","Lucida Sans",Geneva,Verdana,sans-serif'},
    {Name: "Lucida Sans Typewriter", SafeValue: '"Lucida Sans Typewriter","Lucida Console",monaco,"Bitstream Vera Sans Mono",monospace'},
    {Name: "Monaco", SafeValue: 'monaco,Consolas,"Lucida Console",monospace'},
    {Name: "Optima", SafeValue: 'Optima,Segoe,"Segoe UI",Candara,Calibri,Arial,sans-serif'},
    {Name: "Palatino", SafeValue: 'Palatino,"Palatino Linotype","Palatino LT STD","Book Antiqua",Georgia,serif'},
    {Name: "Papyrus", SafeValue: 'Papyrus,fantasy'},
    {Name: "Perpetua", SafeValue: 'Perpetua,Baskerville,"Big Caslon","Palatino Linotype",Palatino,"URW Palladio L","Nimbus Roman No9 L",serif'},
    {Name: "Segoe UI", SafeValue: '"Segoe UI",Frutiger,"Frutiger Linotype","Dejavu Sans","Helvetica Neue",Arial,sans-serif'},
    {Name: "Rockwell", SafeValue: 'Rockwell,"Courier Bold",Courier,Georgia,Times,"Times New Roman",serif'},
    {Name: "Rockwell Extra Bold", SafeValue: '"Rockwell Extra Bold","Rockwell Bold",monospace'},
    {Name: "Tahoma", SafeValue: 'Tahoma,Verdana,Segoe,sans-serif'},
    {Name: "Times New Roman", SafeValue: 'TimesNewRoman,"Times New Roman",Times,Baskerville,Georgia,serif'},
    {Name: "Trebuchet MS", SafeValue: '"Trebuchet MS","Lucida Grande","Lucida Sans Unicode","Lucida Sans",Tahoma,sans-serif'},
    {Name: "Verdana", SafeValue: 'Verdana,Geneva,sans-serif'}
  ];

  private latestValidateValue: string;
  private async: Async;
  private delayedValidate: (value: string) => void;
  private _key: string;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldFontPickerHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.onOpenDialog = this.onOpenDialog.bind(this);
    this.toggleHover = this.toggleHover.bind(this);
    this.toggleHoverLeave = this.toggleHoverLeave.bind(this);
    this.onClickFont = this.onClickFont.bind(this);
    this.onFontDropdownChanged = this.onFontDropdownChanged.bind(this);
    this.mouseEnterDropDown = this.mouseEnterDropDown.bind(this);
    this.mouseLeaveDropDown = this.mouseLeaveDropDown.bind(this);
    this._key = GuidHelper.getGuid();

    //Init the state
    this.state = {
        isOpen: false,
        isHoverDropdown: false,
        errorMessage: ''
      };

    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);

    //Inits the default value
    if (props.initialValue != null && props.initialValue != '') {
      for (var i = 0; i < this.fonts.length; i++) {
        var font = this.fonts[i];
        //Checks if we must use the font name or the font safe value
        if (props.useSafeFont === false && props.initialValue === font.Name) {
          this.state.selectedFont = font.Name;
          this.state.safeSelectedFont = font.SafeValue;
        }
        else if (props.initialValue === font.SafeValue) {
          this.state.selectedFont = font.Name;
          this.state.safeSelectedFont = font.SafeValue;
        }
      }
    }
  }

  /**
   * @function
   * Function to refresh the Web Part properties
   */
  private changeSelectedFont(newValue: string): void {
    this.delayedValidate(newValue);
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

    if (this.latestValidateValue === value)
      return;
    this.latestValidateValue = value;

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || '');
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.initialValue, value);
        this.state.errorMessage = result;
        this.setState(this.state);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.initialValue, value);
          this.state.errorMessage = errorMessage;
          this.setState(this.state);
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
    if (this.props.onPropertyChange && newValue != null) {
      this.props.properties[this.props.targetProperty] = newValue;
      this.props.onPropertyChange(this.props.targetProperty, oldValue, newValue);
    }
  }

  /**
   * @function
   * Called when the component will unmount
   */
  public componentWillUnmount() {
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
    this.state.hoverFont = hoverFont;
    this.setState(this.state);
  }

  /**
   * @function
   * Mouse is leaving a font
   */
  private toggleHoverLeave(element?: any) {
    this.state.hoverFont = '';
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
  private onClickFont(element?: any) {
    var clickedFont: string = element.currentTarget.textContent;
    this.state.selectedFont = clickedFont;
    this.state.safeSelectedFont = this.getSafeFont(clickedFont);
    this.onOpenDialog();
    if (this.props.useSafeFont === false) {
      this.changeSelectedFont(this.state.selectedFont);
    }
    else {
      this.changeSelectedFont(this.state.safeSelectedFont);
    }
  }

  /**
   * @function
   * Gets a safe font value from a font name
   */
  private getSafeFont(fontName: string): string {
    for (var i = 0; i < this.fonts.length; i++) {
      var font = this.fonts[i];
      if (font.Name === fontName)
        return font.SafeValue;
    }
    return '';
  }

  /**
   * @function
   * The font dropdown selected value changed (used when the previewFont property equals false)
   */
  private onFontDropdownChanged(option: IDropdownOption, index?: number): void {
    this.changeSelectedFont(option.key as string);
  }

  /**
   * @function
   * Renders the control
   */
  public render(): JSX.Element {

    if (this.props.previewFonts === false) {
      //If the user don't want to use the preview font picker,
      //we're building a classical drop down picker
      var dropDownOptions: IDropdownOption[] = [];
      var selectedKey: string;
      this.fonts.map((font: ISafeFont) => {
        var isSelected: boolean = false;
        if (this.props.useSafeFont === false && font.Name == this.props.initialValue) {
          isSelected = true;
          selectedKey = font.Name;
        }
        else if (font.SafeValue == this.props.initialValue) {
          isSelected = true;
          selectedKey = font.SafeValue;
        }
        dropDownOptions.push(
          {
            key: this.props.useSafeFont === false ? font.Name : font.SafeValue,
            text: font.Name,
            isSelected: isSelected
          }
        );
      });
      return (
        <div>
          <Dropdown label={this.props.label} options={dropDownOptions} selectedKey={selectedKey}
            onChanged={this.onFontDropdownChanged} disabled={this.props.disabled} />
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
    else {
      //User wants to use the preview font picker, so just build it
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
        fontFamily: this.state.safeSelectedFont != null && this.state.safeSelectedFont != '' ? this.state.safeSelectedFont : 'Arial',
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
        top: '29px',
        left: '0',
        width: 'calc(100% - 2px)',
        //boxShadow: '0 4px 5px rgba(0,0,0,.15)',
        zIndex: 999,
        display: this.state.isOpen ? 'block' : 'none'
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
            <a style={fontSelectA} onClick={this.onOpenDialog}
              onMouseEnter={this.mouseEnterDropDown} onMouseLeave={this.mouseLeaveDropDown} role="menuitem">
              <span style={fontSelectASpan}>{this.state.selectedFont}</span>
              <div style={fontSelectADiv}>
                <i style={fontSelectADivB} className={carret}></i>
              </div>
            </a>
            <div style={fsDrop}>
              <ul style={fsResults}>
                {this.fonts.map((font: ISafeFont, index: number) => {
                  var backgroundColor: string = 'transparent';
                  if (this.state.selectedFont === font.Name)
                    backgroundColor = '#c7e0f4';
                  else if (this.state.hoverFont === font.Name)
                    backgroundColor = '#eaeaea';
                  var innerStyle = {
                    lineHeight: '80%',
                    padding: '7px 7px 8px',
                    margin: '0',
                    listStyle: 'none',
                    fontSize: '18px',
                    fontFamily: font.SafeValue,
                    backgroundColor: backgroundColor,
                    cursor: 'pointer'
                  };
                  return (
                    <li value={font.Name} key={this._key + '-fontpicker-' + index} role="menuitem" onMouseEnter={this.toggleHover} onClick={this.onClickFont} onMouseLeave={this.toggleHoverLeave} style={innerStyle}>{font.Name}</li>
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
}