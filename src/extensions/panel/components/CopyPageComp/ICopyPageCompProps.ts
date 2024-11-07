import { ListViewCommandSetContext, RowAccessor } from "@microsoft/sp-listview-extensibility";
import { SPFI } from "@pnp/sp";
import { IStatefulPanelProps } from "../StatefulPanel/IStatefulPanelProps";

export interface ICopyPageCompProps {
    selectedRows: readonly RowAccessor[];
    spfiContext: SPFI;
    context: ListViewCommandSetContext;
    listName: string;
    listId: string;
    panelConfig: IStatefulPanelProps;
    onCompleted?: (success: boolean) => void;
}