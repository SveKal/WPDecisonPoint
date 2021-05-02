import * as React from "react";
import styles from "./WpDecisionPoint.module.scss";
import { IWpDecisionPointProps } from "./IWpDecisionPointProps";
import { Web, Item } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import IStates from "./IStates";
import ReactTooltip from "react-tooltip";
import { ChoiceGroup, IChoiceGroupOption } from "office-ui-fabric-react";

const options: IChoiceGroupOption[] = [
  { key: "1", text: "1" },
  { key: "2", text: "2" },
  { key: "3", text: "3" },
  { key: "4", text: "4" },
  { key: "5", text: "5" },
  { key: "6", text: "6" },
  { key: "7", text: "7" },
];

export default class WpDecisionPoint extends React.Component<
  IWpDecisionPointProps,
  IStates
> {
  constructor(props) {
    super(props);
    this.state = {
      Item: [],
      HTML: [],
      DecisionPoint: 1,
    };
    //binding function to webcontext
    this._onDecisionPointChange = this._onDecisionPointChange.bind(this);
  }

  public async componentDidMount() {
    await this.fetchDataNewArchitecture();
  }

  //getting data from list/API
  public async fetchDataNewArchitecture() {
    let web = Web(this.props.currentSiteUrl);
    //change List name and change filter if needed
    const item: any[] = await web.lists
      .getByTitle(this.props.listName)
      .items.getById(1)
      .get();
    this.setState({ Item: item });
    let html = await this.getHTML(item);
    //in-web-part setting of status enbled in Edit Mode
    if (document.location.href.indexOf("Mode=Edit") !== -1) {
      document.getElementById("inEditMode").style.display = "inline-block";
    }
    this.setState({ HTML: html });
    console.log(Item);
  }

  //Render Picture, text for current decision point
  public async getHTML(item) {
    var decisionPoint = item.DecisionPoint;
    this.setState({
      DecisionPoint: decisionPoint,
    });

    //Tooltip text
    var toolTipOverDecisionPointPicture = `Beslutspunkter vilka måste genomgås för att 
    <br> få fortsätta till nästkommande processteg. 
    <br> En beslutspunkt innebär alltid en valsituation: 
    <br> "Fortsätta", "Backa för komplettering" eller "Avbryta och lägga ned"`;

    var html = (
      <div>
        <div className={styles.column}>
          <span className={styles.title}>
            Senast passerade beslutspunkt: BP{decisionPoint}{" "}
          </span>
        </div>
        <div className={styles.column}>
          <img
            data-tip={toolTipOverDecisionPointPicture} //setting tooltip text
            className={styles.image}
            src={require("../images/BP" + decisionPoint + ".svg")}
          />
        </div>
        {/* Render Tooltip on mouseover */}
        <ReactTooltip
          multiline
          place="bottom"
          arrowColor="#CC4201"
          backgroundColor="#F18700"
          textColor="black"
          border
          borderColor="black"
        />
      </div>
    );

    return await html;
  }

  //DecisionPoint state update on change of decision point ChoiceGroup
  private _onDecisionPointChange(
    ev: React.FormEvent<HTMLInputElement>,
    option: IChoiceGroupOption
  ) {
    let decisionPoint = Number(option.key);
    this.setState({
      DecisionPoint: decisionPoint,
    });
    this.UpdateStatus(decisionPoint);
  }

  //Updating status list in SharePoint Online
  public async UpdateStatus(decisionPoint) {
    let web = Web(this.props.currentSiteUrl);
    console.log(web);
    await web.lists.getByTitle(this.props.listName).items.getById(1).update({
      DecisionPoint: decisionPoint,
    });
    await this.fetchDataNewArchitecture();
  }

  //Render web part
  public render(): React.ReactElement<IWpDecisionPointProps> {
    return (
      <div className={styles.wpDecisionPoint}>
        <div className={styles.container}>
          <div className={styles.row}>
            {this.state.HTML}
            <div className={styles.columnEdit} id="inEditMode">
              <ChoiceGroup
                className={styles.choice}
                checked
                value={this.state.DecisionPoint}
                selectedKey={this.state.DecisionPoint.toString()}
                options={options}
                onChange={this._onDecisionPointChange}
                label="Uppdatera beslutspunkt"
              />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
