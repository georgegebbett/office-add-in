import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { SignIn, SignInButton, useUser } from "@clerk/clerk-react";

/* global Word, Office, require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default function App({ title, isOfficeInitialized }: AppProps) {
  const [listItems] = React.useState<AppState["listItems"]>([
    {
      icon: "Ribbon",
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: "Unlock",
      primaryText: "Upload documents to your cases!",
    },
    {
      icon: "Design",
      primaryText: "Lawyer like a pro",
    },
  ]);

  const click = async () => {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      await context.sync();
    });
  };

  const { user } = useUser();

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    );
  }

  return (
    <div className="ms-welcome">
      {/* <Header logo={require("./../../../assets/logo-filled.png")} title={title} message="Welcome" /> */}
      <HeroList message="Lawhive Word Add-in" items={listItems}>
        <p className="ms-font-l">Logged in as</p>
        {user ? user.fullName : "Not signed in"}
        <SignIn />
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={click}>
          Run
        </DefaultButton>
      </HeroList>
    </div>
  );
}
