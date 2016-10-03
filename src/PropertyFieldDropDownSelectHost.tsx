/**
 * @file PropertyFieldDropDownSelectHost.tsx
 * Renders the controls for PropertyFieldDropDownSelect component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldDropDownSelectPropsInternal } from './PropertyFieldDropDownSelect';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

/**
 * @interface
 * PropertyFieldDropDownSelectHost properties interface
 *
 */
export interface IPropertyFieldDropDownSelectHostProps extends IPropertyFieldDropDownSelectPropsInternal {
}

/**
 * @interface
 * PropertyFieldDropDownSelectHost state interface
 *
 */
export interface IPropertyFieldDropDownSelectHostState {
  isOpen: boolean;
  isHoverDropdown?: boolean;
  hoverFont?: string;
  selectedFont?: string[];
  safeSelectedFont?: string[];
}

/**
 * @class
 * Renders the controls for PropertyFieldDropDownSelect component
 */
export default class PropertyFieldDropDownSelectHost extends React.Component<IPropertyFieldDropDownSelectHostProps, IPropertyFieldDropDownSelectHostState> {

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldDropDownSelectHostProps) {
    super(props);

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

    if (props.initialValue != null && props.initialValue.length > 0  && this.props.options != null) {
      for (var i = 0; i < this.props.options.length; i++) {
        var font = this.props.options[i];
        var found: boolean = false;
        for (var j = 0; j < props.initialValue.length; j++) {
          if (props.initialValue[j] == font.key) {
            found = true;
            break;
          }
        }
        if (found == true)
          font.isSelected = true;
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
      this.props.onPropertyChange(this.props.targetProperty, newValue);
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

  private saveOptions(): void {
    var res: string[] = [];
    this.props.options.map((elm: IDropdownOption) => {
      if (elm.isSelected)
        res.push(elm.key.toString());
    });
    if (this.props.onPropertyChange && res != null) {
      this.props.onPropertyChange(this.props.targetProperty, res);
    }
  }

  /**
   * @function
   * User clicked on a font
   */
  private onClickFont(element?: any) {
    var clickedFont: string = element.currentTarget.textContent;
    var option: IDropdownOption = this.getOption(clickedFont);
    option.isSelected = !option.isSelected;
    this.setState(this.state);
    this.saveOptions();
  }

  private getOption(text: string): IDropdownOption {
    for (var i = 0; i < this.props.options.length; i++) {
      var font = this.props.options[i];
      if (font.text === text)
        return font;
    }
    return null;
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

      //User wants to use the preview font picker, so just build it
      var fontSelect = {
        fontSize: '16px',
        width: '100%',
        position: 'relative',
        display: 'inline-block',
        zoom: '1'
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
        boxShadow: '0 4px 5px rgba(0,0,0,.15)',
        zIndex: '999',
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
      var foundSelected = false;
      //Renders content
      return (
        <div style={{ marginBottom: '8px'}}>
          <Label>{this.props.label}</Label>
          <div style={fontSelect}>
            <a style={fontSelectA} onClick={this.onOpenDialog}
              onMouseEnter={this.mouseEnterDropDown} onMouseLeave={this.mouseLeaveDropDown}>
              <span style={fontSelectASpan}>
                {this.props.options.map((elm: IDropdownOption, index?: number) => {
                  if (elm.isSelected) {
                    if (foundSelected == false) {
                      foundSelected = true;
                      return (
                          <span>{elm.text}</span>
                      );
                    }
                    else {
                      return (
                          <span>, {elm.text}</span>
                      );
                    }
                  }
                }
                )}
                {this.state.selectedFont}
              </span>
              <div style={fontSelectADiv}>
                <i style={fontSelectADivB} className={carret}></i>
              </div>
            </a>
            <div style={fsDrop}>
              <ul style={fsResults}>
                {this.props.options.map((font: IDropdownOption) => {
                  var backgroundColor: string = 'transparent';
                  if (this.state.hoverFont === font.text)
                    backgroundColor = '#eaeaea';
                  var innerStyle = {
                    lineHeight: '80%',
                    padding: '7px 7px 8px',
                    margin: '0',
                    listStyle: 'none',
                    fontSize: '16px',
                    backgroundColor: backgroundColor,
                    cursor: 'pointer'
                  };
                  return (
                    <li value={font.text} onMouseEnter={this.toggleHover} onClick={this.onClickFont} onMouseLeave={this.toggleHoverLeave} style={innerStyle}>
                      <input style={{width: '18px', height: '18px'}} checked={font.isSelected} type="checkbox" role="checkbox" />
                      {font.text}
                    </li>
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