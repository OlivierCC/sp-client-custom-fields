/**
 * @file PropertyFieldTagPickerHost.tsx
 * Renders the controls for PropertyFieldTagPicker component
 *
 * @copyright 2017 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldTagPickerPropsInternal, IPropertyFieldTag } from './PropertyFieldTagPicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
import { TagPicker, IBasePickerSuggestionsProps, ITag } from 'office-ui-fabric-react/lib/Pickers';

/**
 * @interface
 * PropertyFieldTagPickerHost properties interface
 *
 */
export interface IPropertyFieldTagPickerHostProps extends IPropertyFieldTagPickerPropsInternal {
}

export interface IPropertyFieldTagPickerState {
  errorMessage: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldTagPicker component
 */
export default class PropertyFieldTagPickerHost extends React.Component<IPropertyFieldTagPickerHostProps, IPropertyFieldTagPickerState> {

  private async: Async;
  private delayedValidate: (value: IPropertyFieldTag[]) => void;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldTagPickerHostProps) {
    super(props);

    this.async = new Async(this);
    this.state = { errorMessage: ''};

    //Bind the current object to the external called onSelectDate method
    this.onItemChanged = this.onItemChanged.bind(this);
    this.onFilterChanged = this.onFilterChanged.bind(this);
    this.listContainsTag = this.listContainsTag.bind(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: IPropertyFieldTag[]): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.selectedTags, value);
      return;
    }

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || []);
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.selectedTags, value);
        this.setState({ errorMessage: result} as IPropertyFieldTagPickerState);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.selectedTags, value);
          this.setState({ errorMessage } as IPropertyFieldTagPickerState);
        });
      }
    }
    else {
      this.notifyAfterValidate(this.props.selectedTags, value);
    }
  }

  /**
   * @function
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: IPropertyFieldTag[], newValue: IPropertyFieldTag[]) {
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
   * Called when the TagPicker text changed
   * @param filterText
   * @param tagList
   */
  private onFilterChanged(filterText: string, tagList: ITag[]) {
    return filterText ? this.props.tags.filter(tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0).filter(item => !this.listContainsTag(item, tagList)) : [];
  }

  /**
   * @function
   * Tests if the selected list contains already the tag
   * @param tag
   * @param tagList
   */
  private listContainsTag(tag: ITag, tagList: ITag[]) {
    if (!tagList || !tagList.length || tagList.length === 0) {
      return false;
    }
    return tagList.filter(compareTag => compareTag.key === tag.key).length > 0;
  }

  /**
   * @function
   * Occurs when the list of selected items changed
   * @param selectedItems
   */
  private onItemChanged(selectedItems: ITag[]): void {
    this.delayedValidate(selectedItems);
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
        <TagPicker
          onResolveSuggestions={ this.onFilterChanged }
          getTextFromItem= {(item: ITag) => { return item.name; } }
          defaultSelectedItems={this.props.selectedTags}
          onChange={this.onItemChanged}
          pickerSuggestionsProps={
            {
              suggestionsHeaderText: this.props.suggestionsHeaderText,
              noResultsFoundText: this.props.noResultsFoundText,
              loadingText: this.props.loadingText
            }
          }
        />
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
