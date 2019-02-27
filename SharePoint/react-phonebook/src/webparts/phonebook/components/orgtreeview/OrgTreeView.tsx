import * as React from 'react';
import Tree, {TreeNode} from 'rc-tree'
import 'rc-tree/assets/index.css';
import { IOrgTreeViewProps } from './IOrgTreeViewProps';
import { IOrgTreeViewState } from './IOrgTreeViewState';
import { IOrgUnit } from '../../../../classes/IOrgUnit';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { Guid } from '@microsoft/sp-core-library';





export default class OrgTreeView extends React.Component<IOrgTreeViewProps, IOrgTreeViewState> {
  constructor(props: IOrgTreeViewProps) {
    super(props);

    this.state = {
      isLoading: false,
      units: [],
      selectedUnitId: Guid.empty
    };
    
  }

  public async componentDidMount() {
    this.setState({isLoading:true, units: [], selectedUnitId: Guid.empty});

    let orgunits: IOrgUnit[] = await this.props.getData();
    this.setState({isLoading:false, units: orgunits, selectedUnitId: Guid.empty});
  }

  private generateTreeNodes(node:IOrgUnit){
    let result:TreeNode = {title: node.title, children: [], term_id: node.id.toString()}
    if (node.childOrgunits.length>0){
      node.childOrgunits.forEach((unit:IOrgUnit) => {
        result.children.push(this.generateTreeNodes(unit));
      });
    }
    return result;
  }

  public async componentDidUpdate(prevProps: IOrgTreeViewProps, prevState: IOrgTreeViewState) {
    if (this.props.orgStructureTermSet != prevProps.orgStructureTermSet){
      this.setState({isLoading:true, units: [], selectedUnitId: Guid.empty});

      let orgunits: IOrgUnit[] = await this.props.getData();
      this.setState({isLoading:false, units: orgunits, selectedUnitId: Guid.empty});
    }
  }

  private onSelect(selectedKeys, e:{selected: boolean, selectedNodes, node, event, nativeEvent}) {
    let unitId: Guid = selectedKeys.length > 0 ? Guid.parse(e.node.props.term_id) : Guid.empty;
      this.setState((current) => ({
        isLoading:false, 
        units: current.units, 
        selectedUnitId: unitId
      }));
    this.props.orgUnitSelected(unitId);
  }

  public render(): React.ReactElement<IOrgTreeViewProps> {
    const isLoading: JSX.Element = this.state.isLoading ? <div><Spinner label="Loading" /></div>: <div></div>;

    let treenodes = [];
    this.state.units.forEach((unit:IOrgUnit)=> {
      treenodes.push(this.generateTreeNodes(unit));
    });

    return (
      <div>
        {isLoading}
        <Tree treeData={treenodes} onSelect={this.onSelect.bind(this)}>
        </Tree>
      </div>
    );
  }
}
