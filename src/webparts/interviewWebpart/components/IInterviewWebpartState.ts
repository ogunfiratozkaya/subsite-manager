import SPGroup from "../models/SPGroup";
import User from "../models/User";
import { IColumn } from "office-ui-fabric-react/lib/DetailsList";

export interface IInterviewWebpartState {
    showModal: boolean;
    showPanel: boolean;
    reloadItems: boolean;
    selectedGroup?: SPGroup;
    selectedFieldName?: string;
}