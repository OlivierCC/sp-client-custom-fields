/**
 * @file PropertyFieldSortableListHost.tsx
 * Renders the controls for PropertyFieldSortableList component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 *
 */
import * as React from 'react';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { Spinner, SpinnerType } from 'office-ui-fabric-react/lib/Spinner';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
import GuidHelper from './GuidHelper';
import { IPropertyFieldSortableListPropsInternal, ISortableListOrder } from './PropertyFieldSortableList';

/**
 * @interface
 * PropertyFieldSortableListHost properties interface
 *
 */
export interface IPropertyFieldSortableListHostProps extends IPropertyFieldSortableListPropsInternal {
}

/**
 * @interface
 * PropertyFieldSortableListHost state interface
 *
 */
export interface IPropertyFieldSortableListHostState {
  results: IChoiceGroupOption[];
  selectedKeys: string[];
  errorMessage?: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldSortableList component
 */
export default class PropertyFieldSortableListHost extends React.Component<IPropertyFieldSortableListHostProps, IPropertyFieldSortableListHostState> {

  private async: Async;
  private delayedValidate: (value: string[]) => void;
  private _key: string;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldSortableListHostProps) {
    super(props);

    this._key = GuidHelper.getGuid();
    this.onChanged = this.onChanged.bind(this);
    this.state = {
			results: this.props.items !== undefined ? this.props.items : [],
      selectedKeys: this.props.selectedItems !== undefined ? this.props.selectedItems : [],
      errorMessage: ''
    };

    this.sortDescending = this.sortDescending.bind(this);
    this.sortAscending = this.sortAscending.bind(this);
    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);

    this.initDefaultValue();
  }

  /**
   * Inits the default items checked values
   */
  private initDefaultValue(): void {
    if (this.props.selectedItems !== undefined && this.props.selectedItems != null) {
      for (var i = 0; i < this.props.selectedItems.length; i++) {
        var currItem = this.props.selectedItems[i];
        var choice: IChoiceGroupOption = this.getStateItemFromKey(currItem);
        if (choice != null) {
          choice.isChecked = true;
        }
      }
    }
  }

  /**
   * Gets the item from key
   * @param key
   */
  private getStateItemFromKey(key: string): IChoiceGroupOption {
    for (var i = 0; i < this.state.results.length; i++) {
      var currItem = this.state.results[i];
      if (currItem.key === key)
        return currItem;
    }
    return null;
  }

  /**
   * @function
   * Remove a string from the selected keys
   */
  private removeSelected(element: string): void {
    var res = [];
    for (var i = 0; i < this.state.selectedKeys.length; i++) {
      if (this.state.selectedKeys[i] !== element)
        res.push(this.state.selectedKeys[i]);
    }
    this.state.selectedKeys = res;
    this.getStateItemFromKey(element).isChecked = false;
    this.setState(this.state);
  }

  /**
   * @function
   * Raises when a list has been selected
   */
  private onChanged(element: React.FormEvent<HTMLElement>, isChecked: boolean): void {
    if (element) {
      var value: string = (element.currentTarget as any).value;

      if (isChecked === false) {
        this.removeSelected(value);
      }
      else {
        this.getStateItemFromKey(value).isChecked = true;
        this.state.selectedKeys.push(value);
        this.setState(this.state);
      }
      this.delayedValidate(this.state.selectedKeys);
    }
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: string[]): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.selectedItems, value);
      return;
    }

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || []);
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.selectedItems, value);
        this.state.errorMessage = result;
        this.setState(this.state);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.selectedItems, value);
          this.state.errorMessage = errorMessage;
          this.setState(this.state);
        });
      }
    }
    else {
      this.notifyAfterValidate(this.props.selectedItems, value);
    }
  }

  /**
   * @function
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: string[], newValue: string[]) {
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

  private sortDescending(elm?: any): void {
    this.state.results.sort((a: IChoiceGroupOption, b: IChoiceGroupOption): any => {
      if (this.props.sortBy == ISortableListOrder.Key) {
        return (a.key > b.key) ? 1 : ((b.key > a.key) ? -1 : 0);
      }
      else {
        return (a.text > b.text) ? 1 : ((b.text > a.text) ? -1 : 0);
      }
    });
    this.setState(this.state);
  }

  private sortAscending(elm?: any): void {
    this.state.results.sort((a: IChoiceGroupOption, b: IChoiceGroupOption): any => {
      if (this.props.sortBy == ISortableListOrder.Key) {
        return (a.key > b.key) ? -1 : ((b.key > a.key) ? 1 : 0);
      }
      else {
        return (a.text > b.text) ? -1 : ((b.text > a.text) ? 1 : 0);
      }
    });
    this.setState(this.state);
  }

  /**
   * @function
   * Renders the SPListMultiplePicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {

        var styleOfLabel: any = {
          color: this.props.disabled === true ? '#A6A6A6' : 'auto',
          width: '160px',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap'
        };
        //Renders content
        return (
          <div>
            <Label>{this.props.label}</Label>
            <div style={{position: 'absolute', right: '0', marginRight: '20px', zIndex: 998}}>
              <Button icon="ChevronUp" buttonType={ButtonType.icon} onClick={this.sortDescending} disabled={this.props.disabled}></Button>
              <Button icon="ChevronDown" buttonType={ButtonType.icon} onClick={this.sortAscending} disabled={this.props.disabled}></Button>
            </div>
            {this.state.results.map((item: IChoiceGroupOption, index: number) => {
              var uniqueKey = this.props.targetProperty + '-' + item.key;
              var checked = item.isChecked != null && item.isChecked !== undefined ? item.isChecked : false;
              return (
                <div className="ms-ChoiceField" key={this._key + '-sortablelistpicker-' + index}>
                  <Checkbox
                    checked={checked}
                    disabled={this.props.disabled}
                    label={item.text}
                    onChange={this.onChanged}
                    inputProps={{value: item.key}}
                  />
                </div>
              );
            })
            }
            { this.state.errorMessage != null && this.state.errorMessage != '' && this.state.errorMessage != undefined ?
              <div style={{paddingBottom: '8px'}}><div aria-live='assertive' className='ms-u-screenReaderOnly' data-automation-id='error-message'>{  this.state.errorMessage }</div>
              <span>
                <p className='ms-TextField-errorMessage ms-u-slideDownIn20'>{ this.state.errorMessage }</p>
              </span>
              </div>
            : ''}
          </div>
        );
  }

}
