import { IMail } from "../../../models/IMail";
import { GraphHelper } from "../../../services/GraphHelper";

export interface IOutlookAddinProps {
  description: string;
  mail: IMail;
  graphHelper: GraphHelper;
}
