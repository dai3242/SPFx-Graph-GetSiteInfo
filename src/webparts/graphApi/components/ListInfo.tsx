import * as React from "react";
import {
  Dropdown,
  IDropdownOption,
  IDropdownStyles,
} from "@fluentui/react/lib/Dropdown";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { Context } from "./GraphApi";
import {
  IIconProps,
  IconButton,
  PrimaryButton,
  Text,
  TextField,
} from "@fluentui/react";

type IListData = {
  id: string;
  webUrl: string;
  displayName: string;
  name: string;
};

type IDriveData = {
  id: string;
  webUrl: string;
  name: string;
};

const emojiIcon: IIconProps = { iconName: "Copy" };

const InputSearchTerm = ({
  setSearchTerm,
}: {
  setSearchTerm: React.Dispatch<React.SetStateAction<string | undefined>>;
}) => {
  const handleChange = (
    e: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    value: string | undefined
  ) => {
    setSearchTerm(value);
  };

  return (
    <div style={{ width: 200 }}>
      <TextField
        label="Standard"
        onChange={(e, value) => handleChange(e, value)}
      />
    </div>
  );
};

const SearchItem = ({
  searchTerm,
  driveId,
}: {
  searchTerm: string;
  driveId: string;
}) => {
  const context = React.useContext(Context);

  const getItem = async () => {
    try {
      const graphClient: MSGraphClientV3 | undefined =
        await context?.msGraphClientFactory.getClient("3");

      const Response = await graphClient
        ?.api(`/drives/${driveId}/root/search(q='${searchTerm}')`)
        .get();
      console.log(Response);
    } catch (err) {
      console.log("Error: ", err);
    }
  };

  return (
    <div style={{marginLeft:10}}>
      <PrimaryButton text="Search" onClick={getItem} />
    </div>
  );
};

const SelectList = ({
  selectedList,
  setSelectedList,
  lists,
}: {
  selectedList: IDropdownOption | undefined;
  setSelectedList: React.Dispatch<
    React.SetStateAction<IDropdownOption<any> | undefined>
  >;
  lists: IListData[];
}) => {
  const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };

  const dropdownControlledOptions = lists.map((list) => ({
    key: list.displayName,
    text: list.displayName,
  }));

  const onChange = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    setSelectedList(item);
  };

  return (
    <>
      <Dropdown
        label="Select List"
        selectedKey={selectedList ? selectedList.key : undefined}
        onChange={onChange}
        placeholder="Select List"
        options={dropdownControlledOptions}
        styles={dropdownStyles}
      />
    </>
  );
};

const ListInfo = ({ siteId = "siteId" }: { siteId: string }) => {
  const context = React.useContext(Context);
  const [selectedList, setSelectedList] = React.useState<IDropdownOption>();
  const [lists, setLists] = React.useState<IListData[]>([]);
  const [drives, setDrives] = React.useState<IDriveData[]>([]);
  const [searchTerm, setSearchTerm] = React.useState<string>("");

  React.useEffect(() => {
    const getList = async () => {
      try {
        const graphClient: MSGraphClientV3 | undefined =
          await context?.msGraphClientFactory.getClient("3");

        const Response = await graphClient?.api(`/sites/${siteId}/lists`).get();

        setLists(Response.value);
      } catch (err) {
        console.log("Error: ", err);
      }
    };

    const getDrive = async () => {
      try {
        const graphClient: MSGraphClientV3 | undefined =
          await context?.msGraphClientFactory.getClient("3");

        const Response = await graphClient
          ?.api(`/sites/${siteId}/drives?$select=id,name,webUrl`)
          .get();

        setDrives(Response.value);
      } catch (err) {
        console.log("Error: ", err);
      }
    };

    getList();
    getDrive();
  }, [siteId]);

  console.log("searchTermis", searchTerm);

  const copytext = (text: string): void => {
    navigator.clipboard.writeText(text);
  };

  return (
    <>
      <SelectList
        selectedList={selectedList}
        setSelectedList={setSelectedList}
        lists={lists}
      />
      {lists ? (
        lists
          .filter((list) => list.displayName === selectedList?.text)
          .map((list) => (
            <>
              <div style={{ marginTop: 15 }}>
                <Text variant="large">List Info</Text>
              </div>

              <ul
                style={{ listStyleType: "none", paddingLeft: 10, marginTop: 0 }}
              >
                <li>
                  <Text variant="mediumPlus">
                    Display Name: {list?.displayName}
                  </Text>
                  <IconButton
                    iconProps={emojiIcon}
                    title="Emoji"
                    ariaLabel="Emoji"
                    onClick={() => copytext(list?.displayName)}
                  />
                </li>
                <li>
                  <Text variant="mediumPlus">Name: {list?.name}</Text>
                  <IconButton
                    iconProps={emojiIcon}
                    title="Emoji"
                    ariaLabel="Emoji"
                    onClick={() => copytext(list?.name)}
                  />
                </li>
                <li>
                  <Text variant="mediumPlus">List ID: {list?.id}</Text>
                  <IconButton
                    iconProps={emojiIcon}
                    title="Emoji"
                    ariaLabel="Emoji"
                    onClick={() => copytext(list?.id)}
                  />
                </li>
                <li>
                  <Text variant="mediumPlus">List URL: {list?.webUrl}</Text>
                  <IconButton
                    iconProps={emojiIcon}
                    title="Emoji"
                    ariaLabel="Emoji"
                    onClick={() => copytext(list?.webUrl)}
                  />
                </li>
              </ul>
            </>
          ))
      ) : (
        <div></div>
      )}
      {drives ? (
        drives
          .filter((drive) => drive.name === selectedList?.text)
          .map((drive) => (
            <>
              <div style={{ marginTop: 15 }}>
                <Text variant="large">Drive Info</Text>
              </div>

              <ul
                style={{ listStyleType: "none", paddingLeft: 10, marginTop: 0 }}
              >
                <li>
                  <Text variant="mediumPlus">Drive ID: {drive?.id}</Text>
                  <IconButton
                    iconProps={emojiIcon}
                    title="Emoji"
                    ariaLabel="Emoji"
                    onClick={() => copytext(drive?.id)}
                  />
                </li>
              </ul>
            </>
          ))
      ) : (
        <div></div>
      )}
      <div style={{display:"flex", alignItems:"flex-end"}}>
        <InputSearchTerm setSearchTerm={setSearchTerm} />
        <SearchItem searchTerm={searchTerm} driveId="" />
      </div>
    </>
  );
};

export default ListInfo;
