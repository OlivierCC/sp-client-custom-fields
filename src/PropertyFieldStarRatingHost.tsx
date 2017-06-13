/**
 * @file PropertyFieldStarRatingHost.tsx
 * Renders the controls for PropertyFieldStarRating component
 *
 * @copyright 2017 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldStarRatingPropsInternal } from './PropertyFieldStarRating';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
//import StarRatingComponent from 'react-star-rating-component';
import GuidHelper from './GuidHelper';

var StarRatingComponent: any = require('react-star-rating-component/dist/react-star-rating-component');


/**
 * @interface
 * PropertyFieldStarRatingHost properties interface
 *
 */
export interface IPropertyFieldStarRatingHostProps extends IPropertyFieldStarRatingPropsInternal {
}

export interface IPropertyFieldStarRatingState {
  currentValue?: number;
  errorMessage: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldStarRating component
 */
export default class PropertyFieldStarRatingHost extends React.Component<IPropertyFieldStarRatingHostProps, IPropertyFieldStarRatingState> {

  private async: Async;
  private delayedValidate: (value: number) => void;
  private _key: string;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldStarRatingHostProps) {
    super(props);

    this._key = GuidHelper.getGuid();
    this.async = new Async(this);
    this.state = {
      errorMessage: '',
      currentValue: this.props.initialValue !== undefined ? this.props.initialValue : 0
    };

    //Bind the current object to the external called onSelectDate method
    this.onStarClick = this.onStarClick.bind(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: number): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.initialValue, value);
      return;
    }

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || 0);
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.initialValue, value);
        this.setState({ errorMessage: result} as IPropertyFieldStarRatingState);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.initialValue, value);
          this.setState({ errorMessage } as IPropertyFieldStarRatingState);
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
  private notifyAfterValidate(oldValue: number, newValue: number) {
    this.props.properties[this.props.targetProperty] = newValue;
    this.props.onPropertyChange(this.props.targetProperty, oldValue, newValue);
    if (!this.props.disableReactivePropertyChanges && this.props.render != null)
        this.props.render();
  }

  /**
   * @function
   * Called when the component will unmount
   */
  public componentWillUnmount() {
    this.async.dispose();
  }

  private onStarClick(nextValue: any, prevValue: any, name: string) {
    this.state.currentValue = nextValue;
    this.setState(this.state);
    this.delayedValidate(nextValue);
  }

  /**
   * @function
   * Renders the controls
   */
  public render(): JSX.Element {
    //Renders content
    return (
      <div style={{ marginBottom: '8px'}}>
        <Label>{this.props.label}</Label>
        <div style={{fontSize: this.props.starSize}}>
          <StarRatingComponent
              name={this._key}
              starCount={this.props.starCount}
              starColor={this.props.starColor}
              emptyStarColor={this.props.emptyStarColor}
              value={this.state.currentValue}
              editing={!this.props.disabled}
              onStarClick={this.onStarClick}
              renderStarIcon={null}
              renderStarIconHalf={null}
          />
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
