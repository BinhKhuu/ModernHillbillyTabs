import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneToggle } from "@microsoft/sp-webpart-base";
import {
    CustomCollectionFieldType,
    PropertyFieldCollectionData,
} from "@pnp/spfx-property-controls/lib/PropertyFieldCollectionData";
import * as strings from "HipsterTabsWebPartStrings";
import * as React from "react";
import * as ReactDom from "react-dom";

import HipsterTabs, { IHipsterTabsProps } from "./components/HipsterTabs";
import { IHipsterTab } from "./components/IHipsterTab";
import { IHipsterTabsWebPartProps } from "./IHipsterTabsWebPartProps";

/*
    Changes: v1.0.0.3
        + new pane property spTabId, unique id for a canvas on the page
        + update get zone to get unique canvas id as third prop
        + update moveSection to use unquie canvas ID if sectionID fails
        + add support for selectedTabName url paramter

*/


export default class HipsterTabsWebPart extends BaseClientSideWebPart<IHipsterTabsWebPartProps> {

    protected onInit(): Promise<void> {
        return super.onInit();
    }

    public render(): void {

        /* Code for mapping tabs to a unqiue canvas ID if property tabID is not there
        const zones = this.getZones();
        if(this.properties.tabs.length){
            zones.forEach((zone) => {
                var test = this.properties.tabs.some((tab,index)=>{
                    if( this.properties.tabs[index]["sectionId"] == zone[0]){
                        this.properties.tabs[index]["uniqueId"] = zone[2];
                        return true;
                    }         
                })
            })
        }
        */

        const element: React.ReactElement<IHipsterTabsProps> = React.createElement(
            HipsterTabs,
            {
                instanceId: this.instanceId,
                title: this.properties.title,
                displayMode: this.displayMode,
                updateTitle: (title: string) => {
                    this.properties.title = title;
                },
                configure: () => {
                    this.context.propertyPane.open();
                },
                tabs: this.properties.tabs,
                showAsLinks: this.properties.showAsLinks,
                normalSize: this.properties.normalSize,
            }
        );
        ReactDom.render(element, this.domElement);
    }

    private getZones(): Array<[string, string, string]> {
        const zones = new Array<[string, string, string]>();
        const zoneElements: NodeListOf<HTMLElement> = document.querySelectorAll(".CanvasZoneContainer > .CanvasZone").length > 0 ? <NodeListOf<HTMLElement>>document.querySelectorAll(".CanvasZoneContainer > .CanvasZone") : <NodeListOf<HTMLElement>>document.querySelectorAll("[data-automation-id='CanvasZone']");

        for (let z = 0; z < zoneElements.length; z++) {
            // disqualify the zone containing this webpart
            if (!zoneElements[z].querySelector(`[data-instanceId="${this.instanceId}"]`)) {
                const uniqueId = zoneElements[z].querySelector("[data-automation-id='CanvasControl']").id; //zoneElements[z].querySelector("[data-automation-id='CanvasZone'] [data-automation-id='CanvasControl']")[1].id
                const zoneId = zoneElements[z].dataset.spA11yId;
                const sectionCount = zoneElements[z].getElementsByClassName("CanvasSection").length;
                let zoneName: string = `${strings.PropertyPane_SectionName_Section} ${z + 1} (${sectionCount} ${sectionCount == 1 ? strings.PropertyPane_SectionName_Column : strings.PropertyPane_SectionName_Columns})`;
                zones.push([zoneId, zoneName, uniqueId]);
            }
        }

        return zones;
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
        if (propertyPath == "tabs") {
            // Get Unique tab names
            const tabNames = new Array<string>();
            this.properties.tabs.forEach((tab: IHipsterTab) => {
                if (tabNames.indexOf(tab.name) == -1) {
                    tabNames.push(tab.name);
                }
            });

            // Group entries by tab name (preserving the order)
            // also removes duplicate section entries
            const groupedTabs = new Array<IHipsterTab>();
            const assignedSections = new Array<string>();
            tabNames.forEach((name: string) => {
                groupedTabs.push(...
                    this.properties.tabs.filter((tab: IHipsterTab) => {
                        if (tab.name == name) {
                            if (assignedSections.indexOf(tab.sectionId) == -1) {
                                assignedSections.push(tab.sectionId);
                                return true;
                            }
                        }
                        return false;
                    })
                );
            });

            this.properties.tabs = groupedTabs;
        }
    }



    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    groups: [
                        {
                            groupFields: [
                                PropertyFieldCollectionData("tabs", {
                                    key: "tabs",
                                    label: strings.PropertyPane_TabsLabel,
                                    panelHeader: strings.PropertyPane_TabsHeader,
                                    manageBtnLabel: strings.PropertyPane_TabsButtonLabel,
                                    value: this.properties.tabs,

                                    fields: [
                                        {
                                            id: "name",
                                            title: strings.PropertyPane_TabsField_Name,
                                            type: CustomCollectionFieldType.string,
                                            required: true,
                                        },
                                        {
                                            id: "sectionId",
                                            title: strings.PropertyPane_TabsField_Section,
                                            type: CustomCollectionFieldType.dropdown,
                                            required: true,
                                            options: this.getZones().map((zone: [string, string, string]) => {
                                                return {
                                                    key: zone["0"],
                                                    text: zone["1"],
                                                };
                                            }),
                                        },
                                        {
                                            id: "spTabId",
                                            title: "tabId",
                                            type: CustomCollectionFieldType.custom,
                                            required: true,
                                            disableEdit: true,
                                            onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                                                this.getZones().map((zone: [string, string, string]) => {
                                                    if (item.sectionId == zone[0]) {
                                                        value = zone[2];
                                                        item.spTabId = zone[2];
                                                    }
                                                })
                                                return (
                                                    React.createElement("div", null,
                                                        React.createElement("input", {
                                                            disabled: true, key: value, value: value, onChange: (event: React.FormEvent<HTMLInputElement>) => {
                                                                onUpdate(field.id, event.currentTarget.value);
                                                                if (event.currentTarget.value === "error") {
                                                                    onError(field.id, "Value shouldn't be equal to error");
                                                                } else {
                                                                    onError(field.id, "");
                                                                }
                                                            }
                                                        }),
                                                    )
                                                );
                                            }
                                        },
                                    ],
                                }),
                                PropertyPaneToggle("showAsLinks", {
                                    label: strings.PropertyPane_LinksLabel,
                                    checked: this.properties.showAsLinks,
                                    onText: strings.PropertyPane_LinksOnLabel,
                                    offText: strings.PropertyPane_LinksOffLabel,
                                }),
                                PropertyPaneToggle("normalSize", {
                                    label: strings.PropertyPane_SizeLabel,
                                    checked: this.properties.normalSize,
                                    onText: strings.PropertyPane_SizeOnLabel,
                                    offText: strings.PropertyPane_SizeOffLabel,
                                }),
                            ]
                        },
                    ]
                }
            ]
        };
    }
}
