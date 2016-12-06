/**
 * @file PropertyFieldFontSizePickerHost.tsx
 * Renders the controls for PropertyFieldFontSizePicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldFontSizePickerPropsInternal } from './PropertyFieldFontSizePicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

/**
 * @interface
 * PropertyFieldFontSizePickerHost properties interface
 *
 */
export interface IPropertyFieldFontSizePickerHostProps extends IPropertyFieldFontSizePickerPropsInternal {
}

/**
 * @interface
 * PropertyFieldFontSizePickerHost state interface
 *
 */
export interface IPropertyFieldFontSizePickerHostState {
  isOpen: boolean;
  isHoverDropdown?: boolean;
  hoverFont?: string;
  selectedFont?: string;
  safeSelectedFont?: string;
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
 * Renders the controls for PropertyFieldFontSizePicker component
 */
export default class PropertyFieldFontSizePickerHost extends React.Component<IPropertyFieldFontSizePickerHostProps, IPropertyFieldFontSizePickerHostState> {

  private fonts: ISafeFont[];
  /**
   * @var
   * Defines the font series
   */
  private fontsPixels: ISafeFont[] = [
    {Name: "8px", SafeValue: '8px'},
    {Name: "9px", SafeValue: '9px'},
    {Name: "10px", SafeValue: '10px'},
    {Name: "11px", SafeValue: '11px'},
    {Name: "12px", SafeValue: '12px'},
    {Name: "14px", SafeValue: '14px'},
    {Name: "16px", SafeValue: '16px'},
    {Name: "18px", SafeValue: '18px'},
    {Name: "20px", SafeValue: '20px'},
    {Name: "24px", SafeValue: '24px'},
    {Name: "28px", SafeValue: '28px'},
    {Name: "36px", SafeValue: '36px'},
    {Name: "48px", SafeValue: '48px'}
  ];
  private fontsLabels: ISafeFont[] = [
    {Name: "xx-small", SafeValue: 'xx-small'},
    {Name: "x-small", SafeValue: 'x-small'},
    {Name: "small", SafeValue: 'small'},
    {Name: "medium", SafeValue: 'medium'},
    {Name: "large", SafeValue: 'large'},
    {Name: "x-large", SafeValue: 'x-large'},
    {Name: "xx-large", SafeValue: 'xx-large'}
  ];

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldFontSizePickerHostProps) {
    super(props);

    if (props.usePixels === true)
      this.fonts = this.fontsPixels;
    else
      this.fonts = this.fontsLabels;

    //Bind the current object to the external called onSelectDate method
    this.onOpenDialog = this.onOpenDialog.bind(this);
    this.toggleHover = this.toggleHover.bind(this);
    this.toggleHoverLeave = this.toggleHoverLeave.bind(this);
    this.onClickFont = this.onClickFont.bind(this);
    this.onFontDropdownChanged = this.onFontDropdownChanged.bind(this);
    this.mouseEnterDropDown = this.mouseEnterDropDown.bind(this);
    this.mouseLeaveDropDown = this.mouseLeaveDropDown.bind(this);

    //Init the state
    this.state = {
        isOpen: false,
        isHoverDropdown: false
      };

    //Inits the default value
    if (props.initialValue != null && props.initialValue != '') {
      for (var i = 0; i < this.fonts.length; i++) {
        var font = this.fonts[i];
        //Checks if we must use the font name or the font safe value
        if (props.usePixels === false && props.initialValue === font.Name) {
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
    //Checks if there is a method to called
    if (this.props.onPropertyChange && newValue != null) {
      this.props.properties[this.props.targetProperty] = newValue;
      this.props.onPropertyChange(this.props.targetProperty, this.props.initialValue, newValue);
    }
  }

  /**
   * @function
   * Function to open the dialog
   */
  private onOpenDialog(): void {
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
    if (this.props.usePixels === false) {
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
   * Renders the datepicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {

    if (this.props.preview === false) {
      //If the user don't want to use the preview font picker,
      //we're building a classical drop down picker
      var dropDownOptions: IDropdownOption[] = [];
      var selectedKey: string;
      this.fonts.map((font: ISafeFont) => {
        var isSelected: boolean = false;
        if (this.props.usePixels === false && font.Name == this.props.initialValue) {
          isSelected = true;
          selectedKey = font.Name;
        }
        else if (font.SafeValue == this.props.initialValue) {
          isSelected = true;
          selectedKey = font.SafeValue;
        }
        dropDownOptions.push(
          {
            key: this.props.usePixels === false ? font.Name : font.SafeValue,
            text: font.Name,
            isSelected: isSelected
          }
        );
      });
      return (
        <Dropdown label={this.props.label} options={dropDownOptions} selectedKey={selectedKey}
          onChanged={this.onFontDropdownChanged} />
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
      if (this.state.isOpen === true)
        dropdownColor = '1px solid #3091DE';
      else if (this.state.isHoverDropdown === true)
        dropdownColor = '1px solid #767676';
      var fontSelectA = {
        backgroundColor: '#fff',
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
        color: '#444',
        textDecoration: 'none',
        cursor: 'pointer'
      };
      var fontSelectASpan = {
        marginRight: '26px',
        display: 'block',
        overflow: 'hidden',
        whiteSpace: 'nowrap',
        lineHeight: '1.8',
        textOverflow: 'ellipsis',
        cursor: 'pointer',
        //fontFamily: this.state.safeSelectedFont != null && this.state.safeSelectedFont != '' ? this.state.safeSelectedFont : 'Arial',
        //fontSize: this.state.safeSelectedFont,
        fontWeight: '400'
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
        cursor: 'pointer',
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
                {this.fonts.map((font: ISafeFont) => {
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
                    fontSize: font.SafeValue,
                    backgroundColor: backgroundColor,
                    cursor: 'pointer'
                  };
                  return (
                    <li value={font.Name} role="menuitem" onMouseEnter={this.toggleHover} onClick={this.onClickFont} onMouseLeave={this.toggleHoverLeave} style={innerStyle}>{font.Name}</li>
                  );
                })
                }
              </ul>
            </div>
          </div>
        </div>
      );
    }
  }
}