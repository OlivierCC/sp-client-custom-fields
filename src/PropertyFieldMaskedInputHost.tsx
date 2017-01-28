/**
 * @file PropertyFieldMaskedInputHost.tsx
 * Renders the controls for PropertyFieldMaskedInput component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldMaskedInputPropsInternal } from './PropertyFieldMaskedInput';
import { Label } from 'office-ui-fabric-react/lib/Label';

/**
 * @interface
 * PropertyFieldMaskedInputHost properties interface
 *
 */
export interface IPropertyFieldMaskedInputHostProps extends IPropertyFieldMaskedInputPropsInternal {
}

/**
 * @interface
 * Defines the masked input properties
 *
 */
interface IMaskedInputProps {
  type?: string;
  id?: string;
  placeholder?: string;
  className?: string;
  pattern?: string;
  maxLength?: string;
  title?: string;
  label?: string;
  dataCharset?: string;
  required?: boolean;
  value?: string;
  initialValue?: string;
  onChange?(newValue?: string): void;
  disabled?: boolean;
}

/**
 * @interface
 * Defines the state of masked input control
 *
 */
interface IMaskedInputState {
  firstLoading?: boolean;
  value?: string;
}

/**
 * @interface
 * MaskedInput control.
 * This control is a fork of the input masking component available on GitHub
 * https://github.com/estelle/input-masking
 * by Estelle Weyl. & Alex Schmitz (c)
 *
 */
class MaskedInput extends React.Component<IMaskedInputProps, IMaskedInputState> {

  constructor(props: IPropertyFieldMaskedInputHostProps) {
    super(props);
    //Binds events
    this.handleChange = this.handleChange.bind(this);
    this.handleFocus = this.handleFocus.bind(this);
    this.handleBlur = this.handleBlur.bind(this);

    //Inits default value
    this.state = {
      firstLoading: true,
      value: this.props.initialValue != null ? this.props.initialValue : ''
    };

  }

  public componentDidMount(): void {
    var e = this.refs['inputShell'];
    var event = { target: e };
    this.handleChange(event);
  }

  private handleChange(e): void {
    var previousValue = this.state.value;
    e.target.value = this.handleCurrentValue(e);
    this.state.value = e.target.value;

    if (this.state.firstLoading === true && previousValue == '')
      this.state.value = '';
    this.state.firstLoading = false;
    this.setState(this.state);

    //var maskElement = document.getElementById(this.props.id + 'Mask');
    //if (maskElement != null)
    //  maskElement.innerHTML = this.setValueOfMask(e);

    if (this.props.onChange != null)
      this.props.onChange(e.target.value);
  }

  private handleCurrentValue(e): string {
    var isCharsetPresent = e.target.getAttribute('data-charset'),
        maskedNumber = 'XMDY',
        maskedLetter = '_',
        placeholder = isCharsetPresent || e.target.getAttribute('data-placeholder'),
        value = e.target.value, l = placeholder.length, newValue = '',
        i, j, isInt, isLetter, strippedValue, matchesNumber, matchesLetter;

    // strip special characters
    strippedValue = isCharsetPresent ? value.replace(/\W/g, "") : value.replace(/\D/g, "");

    for (i = 0, j = 0; i < l; i++) {
        isInt = !isNaN(parseInt(strippedValue[j]));
        isLetter = strippedValue[j] ? strippedValue[j].match(/[A-Z]/i) : false;
        matchesNumber = (maskedNumber.indexOf(placeholder[i]) >= 0);
        matchesLetter = (maskedLetter.indexOf(placeholder[i]) >= 0);
        if ((matchesNumber && isInt) || (isCharsetPresent && matchesLetter && isLetter)) {
                newValue += strippedValue[j++];
          } else if ((!isCharsetPresent && !isInt && matchesNumber) || (isCharsetPresent && ((matchesLetter && !isLetter) || (matchesNumber && !isInt)))) {
                //this.options.onError( e ); // write your own error handling function
                return newValue;
        } else {
            newValue += placeholder[i];
        }
        // break if no characters left and the pattern is non-special character
        if (strippedValue[j] == undefined) {
          break;
        }
    }

    if (this.props['data-valid-example']) {
      return this.validateProgress(e, newValue);
    }

    return newValue;
  };

/*
  private setValueOfMask(e): string {
    var value = e.target.value,
        placeholder = e.target.getAttribute('data-placeholder');

    return "<i style=\"font-style: normal;color: transparent;opacity: 0;visibility: hidden;\">" + value + "</i>" + placeholder.substr(value.length);
  };
*/

  private validateProgress(e, value): string {
      var validExample = this.props['data-valid-example'],
          pattern = new RegExp(this.props.pattern),
          placeholder = e.target.getAttribute('data-placeholder'),
          l = value.length, testValue = '', i;

      //convert to months
      if ((l == 1) && (placeholder.toUpperCase().substr(0,2) == 'MM')) {
        if(value > 1 && value < 10) {
          value = '0' + value;
        }
        return value;
      }

    for ( i = l; i >= 0; i--) {
        testValue = value + validExample.substr(value.length);
        if (pattern.test(testValue)) {
          return value;
        } else {
          value = value.substr(0, value.length-1);
        }
    }

      return value;
  };

  private handleBlur(e): void {
      var currValue = e.target.value, pattern;

      // if value is empty, remove label parent class
      if(currValue.length == 0) {

        if(e.target.required) {
          this.updateLabelClass(e, "required", true);
          this.handleError(e, 'required');
        }

      } else {
        pattern = new RegExp('^' + this.props.pattern + '$');

        if(pattern.test(currValue)) {
          this.updateLabelClass(e, "good", true);
        } else {
          this.updateLabelClass(e, "error", true);
          this.handleError(e, 'invalidValue');
        }

      }
  };

  private handleFocus(e): void {
        this.updateLabelClass(e, 'focus', false);
  };

  private updateLabelClass(e, className, replaceExistingClass): void {
       var parentLI = e.target.parentNode.parentNode,
           pastClasses = ['error', 'required', 'focus', 'good'],
           i;

       if (replaceExistingClass) {
           for(i = 0; i < pastClasses.length; i++) {
                parentLI.classList.remove(pastClasses[i]);
           }
       }

       parentLI.classList.add(className);
  };

  private handleError(e, errorMsg): boolean {
        // the event and errorMsg name are passed. Label is already handled. What else do we do with error?
        //var possibleErrorMsgs = ['invalidValue', 'required'];
        return true;
  };

  public render(): JSX.Element {
        var props = {
                 type: (this.props && this.props.type) || '' ,
                 id: this.props.id,
                 placeholder: this.props.placeholder,
                 className: "masked " + (this.props.className || ''),
                 pattern: this.props.pattern,
                 maxLength: this.props.pattern.length,
                 title: this.props.title,
                 label: this.props.label,
                 dataCharset: this.props['data-charset'],
                 required: this.props.required,
                 initialValue: this.props.initialValue,
                 disabled: this.props.disabled
             };

        var shellStyle = {
          position: 'relative',
          lineHeight: '1',
        };
        var shellStyleSpan = {
            position: 'absolute',
            left: '12px',
            top: '3px',
            color: '#ccc',
            pointerEvents: 'none',
            fontSize: '16px',
            fontFamily: 'monospace',
            paddingRight: '10px',
            backgroundColor: 'transparent',
            textTransform: 'uppercase'
        };
        var shellStyleSpanI = {
              fontStyle: 'normal',
              color: 'transparent',
              //opacity: '0',
              visibility: 'hidden'
        };
        var inputShell = {
            fontSize: '16px',
            fontFamily: 'monospace',
            paddingRight: '10px',
            backgroundColor: 'transparent',
            textTransform: 'uppercase'
        };

      var placeHolderContent = props.placeholder.substr(this.state.value.length);

      return (
            <span style={shellStyle}>
                <span style={shellStyleSpan}
                  aria-hidden="true"
                  ref="spanMask"
                  id={props.id + 'Mask'}><i style={shellStyleSpanI}>{this.state.value}</i>{placeHolderContent}</span>
                <input style={inputShell}
                id={props.id}
                ref="inputShell"
                disabled={props.disabled}
                onChange={this.handleChange}
                onFocus={this.handleFocus}
                onBlur={this.handleBlur}
                name={props.id}
                //type={props.type}
                className={props.className}
                data-placeholder={props.placeholder}
                data-pattern={props.pattern}
                aria-required={props.required}
                data-charset={props.dataCharset}
                required={props.required}
                value={this.state.value}
                title={props.title}/>
              </span>
          );
      };
}

/**
 * @class
 * Renders the controls for PropertyFieldMaskedInput component
 */
export default class PropertyFieldMaskedInputHost extends React.Component<IPropertyFieldMaskedInputHostProps, {}> {

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldMaskedInputHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.onValueChanged = this.onValueChanged.bind(this);
  }

  /**
   * @function
   * Function called when the the text changed
   */
  private onValueChanged(element: string): void {
    //Checks if there is a method to called
    if (this.props.onPropertyChange && element != null) {
      this.props.properties[this.props.targetProperty] = element;
      this.props.onPropertyChange(this.props.targetProperty, this.props.initialValue, element);
    }
  }

  /**
   * @function
   * Renders the datepicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {

    //Renders content
    return (
      <div style={{ marginBottom: '8px'}}>
        <Label>{this.props.label}</Label>
        <MaskedInput
            id="tel"
            type="tel"
            disabled={this.props.disabled}
            placeholder={this.props.placeholder}
            pattern={this.props.pattern}
            className="ms-TextField-field"
            maxLength={this.props.maxLength}
            onChange={this.onValueChanged}
            initialValue={this.props.initialValue}
        />
      </div>
    );
  }
}