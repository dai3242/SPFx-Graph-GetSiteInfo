import * as React from "react";
import { IGraphApiProps, IContextProps } from "./IGraphApiProps";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ILabelStyles, Label, Pivot, PivotItem } from "@fluentui/react";
import SiteInfo from "./SiteInfo";
import ListInfo from "./ListInfo";

export const Context = React.createContext<WebPartContext | undefined>(
  undefined
);
export const Properties = React.createContext<IGraphApiProps | undefined>(
  undefined
);

const labelStyles: Partial<ILabelStyles> = {
  root: { marginTop: 10 },
};

const GraphApi: React.FunctionComponent<IContextProps> = (props) => {
  const { context, properties } = props;
  const [ siteId, setSiteId ] = React.useState<string>("");

  return (
    <Context.Provider value={context}>
      <Properties.Provider value={properties}>
        <Pivot aria-label="Basic Pivot Example">
          <PivotItem headerText="Site Info">
            <Label styles={labelStyles}>
              <SiteInfo setSiteId={setSiteId}/>
            </Label>
          </PivotItem>
          <PivotItem headerText="List Info">
            <Label styles={labelStyles}>
              <ListInfo siteId={siteId}/>
            </Label>
          </PivotItem>
        </Pivot>
      </Properties.Provider>
    </Context.Provider>
  );
};

export default GraphApi;
