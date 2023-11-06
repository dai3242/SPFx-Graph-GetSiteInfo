import * as React from "react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { IIconProps, IconButton, Shimmer, Text } from "@fluentui/react";
import { Context } from "./GraphApi";

type ISiteCollectionData =
  | {
      id: string;
      webUrl: string;
      displayName: string;
    }
  | undefined;

const emojiIcon: IIconProps = { iconName: "Copy" };

const SiteInfo = ({
  setSiteId,
}: {
  setSiteId: React.Dispatch<React.SetStateAction<string>>;
}) => {
  const [siteCollectionData, setSiteCollectionData] =
    React.useState<ISiteCollectionData>(undefined);

  const context = React.useContext(Context);

  React.useEffect(() => {
    const getSiteInfo = async () => {
      try {
        const graphClient: MSGraphClientV3 | undefined =
          await context?.msGraphClientFactory.getClient("3");

        const siteCollectionResponse = await graphClient
          ?.api(`/sites?search=${context?.pageContext.web.title}`)
          .get();

        setSiteCollectionData(siteCollectionResponse.value[0]);
        setSiteId(siteCollectionResponse.value[0]?.id);
      } catch (err) {
        console.log("Error: ", err);
      }
    };

    getSiteInfo();
  }, []);

  const copytext = (text: string): void => {
    navigator.clipboard.writeText(text);
  };

  return (
    <>
      {/* <DefaultButton text="Display this site info" onClick={getSiteInfo} /> */}
      {siteCollectionData ? (
        <>
          <ul
            style={{
              listStyleType: "none",
              paddingTop: 0,
              paddingLeft: 10,
              marginTop: 0,
            }}
          >
            <li>
              <Text variant="mediumPlus">
                Display Name: {siteCollectionData?.displayName}
              </Text>
              <IconButton
                iconProps={emojiIcon}
                title="Emoji"
                ariaLabel="Emoji"
                onClick={() => copytext(siteCollectionData?.displayName)}
              />
            </li>
            <li>
              <Text variant="mediumPlus">
                Site ID: {siteCollectionData?.id}
              </Text>
              <IconButton
                iconProps={emojiIcon}
                title="Emoji"
                ariaLabel="Emoji"
                onClick={() => copytext(siteCollectionData?.id)}
              />
            </li>
            <li>
              <Text variant="mediumPlus">
                Site URL: {siteCollectionData?.webUrl}
              </Text>
              <IconButton
                iconProps={emojiIcon}
                title="Emoji"
                ariaLabel="Emoji"
                onClick={() => copytext(siteCollectionData?.webUrl)}
              />
            </li>
          </ul>
        </>
      ) : (
        <>
          <Shimmer width="75%" style={{ marginTop: 10 }} />
          <Shimmer width="75%" style={{ marginTop: 10 }} />
          <Shimmer width="75%" style={{ marginTop: 10 }} />
          <Shimmer width="75%" style={{ marginTop: 10 }} />
        </>
      )}
    </>
  );
};

export default SiteInfo;
