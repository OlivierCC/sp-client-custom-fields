/**
 * @file PropertyFieldTreeViewHost.tsx
 * Renders the controls for PropertyFieldTreeView component
 *
 * @copyright 2017 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldTreeViewPropsInternal, ITreeViewNode } from './PropertyFieldTreeView';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';

require('react-ui-tree-draggable/dist/react-ui-tree.css');
var Tree: any = require('react-ui-tree-draggable/dist/react-ui-tree');

/**
 * @interface
 * PropertyFieldTreeViewHost properties interface
 *
 */
export interface IPropertyFieldTreeViewHostProps extends IPropertyFieldTreeViewPropsInternal {
}

export interface IPropertyFieldTreeViewState {
  errorMessage: string;
  tree: ITreeViewNode[];
  activeNodes: ITreeViewNode[];
}

/**
 * @class
 * Renders the controls for PropertyFieldTreeView component
 */
export default class PropertyFieldTreeViewHost extends React.Component<IPropertyFieldTreeViewHostProps, IPropertyFieldTreeViewState> {

  private async: Async;
  private delayedValidate: (value: string[]) => void;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldTreeViewHostProps) {
    super(props);

    this.async = new Async(this);
    this.state = ({
      errorMessage: '',
      tree: this.props.tree,
      activeNodes: this.getDefaultActiveNodesFromTree()
    } as IPropertyFieldTreeViewState);

    this.renderNode = this.renderNode.bind(this);
    this.onClickNode = this.onClickNode.bind(this);
    this.saveSelectedNodes = this.saveSelectedNodes.bind(this);
    this.handleTreeChange = this.handleTreeChange.bind(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
  }

  private getDefaultActiveNodesFromTree(): ITreeViewNode[] {
    var res: ITreeViewNode[] = [];
    this.props.tree.map((node: ITreeViewNode) => {
      var subTreeViewNodes: ITreeViewNode[] = this.getDefaultActiveNodes(node);
      subTreeViewNodes.map((subNode: ITreeViewNode) => {
        res.push(subNode);
      });
    });
    return res;
  }

  /**
   * @function
   * Gets the list of activated nodes from the  selectedNodesIDs property
   * @param node
   */
  private getDefaultActiveNodes(node: ITreeViewNode): ITreeViewNode[] {
    var res: ITreeViewNode[] = [];
    if (node === undefined || node == null || this.props.selectedNodesIDs === undefined || this.props.selectedNodesIDs == null)
      return res;
    if (this.props.selectedNodesIDs.indexOf(node.id) != -1)
      res.push(node);
    if (node.children !== undefined) {
      for (var i = 0; i < node.children.length; i++) {
        var subTreeViewNodes: ITreeViewNode[] = this.getDefaultActiveNodes(node.children[i]);
        subTreeViewNodes.map((subNode: ITreeViewNode) => {
          res.push(subNode);
        });
      }
    }
    return res;
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: string[]): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.selectedNodesIDs, value);
      return;
    }

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || []);
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.selectedNodesIDs, value);
        this.state.errorMessage = result;
        this.setState(this.state);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.selectedNodesIDs, value);
          this.state.errorMessage = errorMessage;
          this.setState(this.state);
        });
      }
    }
    else {
      this.notifyAfterValidate(this.props.selectedNodesIDs, value);
    }
  }

  /**
   * @function
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: string[], newValue: string[]) {
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
   * Gets the given node position in the active nodes collection
   * @param node
   */
  private getSelectedNodePosition(node: ITreeViewNode): number {
    for (var i = 0; i < this.state.activeNodes.length; i++) {
      if (node === this.state.activeNodes[i])
        return i;
    }
    return -1;
  }

  /**
   * @function
   * Renders the given node
   * @param node
   */
  private renderNode(node: ITreeViewNode): JSX.Element {
    var style: any = { padding: '4px 5px', width: '100%', display: 'flex'};
    var selected: boolean = this.getSelectedNodePosition(node) != -1;
    if (selected === true) {
      style.backgroundColor = '#EAEAEA';
    }
    var isFolder: boolean = false;
    if (node.leaf === false || (node.children !== undefined && node.children.length != 0))
      isFolder = true;
    var checkBoxAvailable: boolean = this.props.checkboxEnabled;
    if (this.props.allowFoldersSelections === false && isFolder === true)
      checkBoxAvailable = false;
    var picUrl: string = '';
    if (selected === true && node.selectedPictureUrl !== undefined)
      picUrl = node.selectedPictureUrl;
    else if (node.collapsed !== true && node.expandedPictureUrl !== undefined)
      picUrl = node.expandedPictureUrl;
    else if (node.pictureUrl !== undefined)
      picUrl = node.pictureUrl;
    return (
        <div style={style} onClick={this.onClickNode.bind(null, node)} name={node.id} id={node.id} role="menuitem">
          { checkBoxAvailable ?
               <div style={{marginRight: '5px'}}> <Checkbox
                    checked={selected}
                    disabled={this.props.disabled}
                    label=''
                    onChange={this.onClickNode.bind(null, node)}
                  />
                </div>
            : ''
          }
          <div style={{paddingTop: '7px'}}>
          {
            picUrl !== undefined && picUrl != '' ?
              <img src={picUrl} width="18" height="18" style={{paddingRight: '5px'}} alt={node.label}/>
            : ''
          }
          {node.label}
          </div>
        </div>
    );
  }

  /**
   * clicks on a node
   * @param node
   */
  private onClickNode(node: ITreeViewNode): void {
    if (this.props.allowFoldersSelections === false && (node.children !== undefined && node.children.length != 0))
      return;
    if (this.props.allowMultipleSelections === false) {
      this.state.activeNodes = [node];
    }
    else {
      var index = this.getSelectedNodePosition(node);
      if (index != -1)
        this.state.activeNodes.splice(index, 1);
      else
        this.state.activeNodes.push(node);
    }
    this.setState(this.state);
    this.saveSelectedNodes();
  }

  /**
   * Saves the selected nodes
   */
  private saveSelectedNodes(): void {
    var res: string[] = [];
    for (var i = 0; i < this.state.activeNodes.length; i++) {
      res.push(this.state.activeNodes[i].id);
    }
    this.delayedValidate(res);
  }

  /**
   * Handles tree changes
   * @param rootNode
   * @param index
   */
  private handleTreeChange(rootNode: any, index: number): void {
    this.state.tree[index] = rootNode;
    this.setState(this.state);
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
        { this.state.tree.map((rootNode: ITreeViewNode, index: number) => {
            return (
              <Tree
                paddingLeft={this.props.nodesPaddingLeft}
                tree={rootNode}
                isNodeCollapsed={false}
                onChange={this.handleTreeChange.bind(null, rootNode, index)}
                renderNode={this.renderNode}
                draggable={false}
                key={'rootNode-' + index}
              />
            );
          })
        }

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
