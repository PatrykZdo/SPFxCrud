import { IDropdownOption } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICrudTestProps {
  spcontext: WebPartContext,
  choices: IDropdownOption[],
}
