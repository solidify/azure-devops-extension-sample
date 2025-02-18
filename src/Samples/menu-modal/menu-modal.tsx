import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";

import { Page } from "azure-devops-ui/Page";

import { showRootComponent } from "../../Common";
import { CommonServiceIds, IProjectPageService, getClient } from "azure-devops-extension-api";
import { WorkItemTrackingRestClient,  } from "azure-devops-extension-api/WorkItemTracking";
import { Button } from "azure-devops-ui/Button";
import { TextField } from "azure-devops-ui/TextField";

interface IMenuModal {
    workItems: any[];
    selectedTags: string[];
    selectedRow: number | null;
    newTag: string;
}

class MenuModal extends React.Component<{}, IMenuModal> {

    constructor(props: {}) {
        super(props);
        this.state = { workItems: [], selectedTags: [], selectedRow: null, newTag: "" };
    }

    public componentDidMount() {
        try {
            console.log("Component did mount, initializing SDK...");
            SDK.init();

            SDK.ready().then(() => {
                console.log("SDK is ready, loading project context and workitems...");
                this.loadProjectContext();
                SDK.resize(400, 400);
            }).catch((error) => {
                console.error("SDK ready failed: ", error);
            });
        } catch (error) {
            console.error("Error during SDK initialization or project context loading: ", error);
        }
    }

    private handleRowClick = (item: any) => {
        const tags = item.fields["System.Tags"] ? item.fields["System.Tags"].split("; ") : [];
        this.setState({ selectedTags: tags, selectedRow: item.id });
    };

    private updateTags = async () => {
        const { selectedRow, selectedTags } = this.state;
        if (selectedRow !== null) {
            try {
                const client = getClient(WorkItemTrackingRestClient);
                await client.updateWorkItem(
                    [
                        {
                            op: "replace",
                            path: "/fields/System.Tags",
                            value: selectedTags.join("; ")
                        }
                    ],
                    selectedRow,

                );
                this.loadProjectContext();
                this.setState({ selectedRow: null, selectedTags: [], newTag: "" });
            } catch (error) {
                console.error("Failed to update tags: ", error);
            }
        }
    };

    private addTag = () => {
        const { newTag, selectedTags } = this.state;
        if (newTag && !selectedTags.includes(newTag)) {
            this.setState({ selectedTags: [...selectedTags, newTag], newTag: "" });
        }
    };

    private removeTag = (tagToRemove: string) => {
        this.setState({ selectedTags: this.state.selectedTags.filter(tag => tag !== tagToRemove) });
    };

    public render(): JSX.Element {
        return (
            <Page>
                <div style={{ flex: 1, overflow: "auto" }}>
                    <h4>Work Items:</h4>
                    <table style={{ width: "100%", tableLayout: "auto", borderCollapse: "collapse" }}>
                        <thead>
                            <tr>
                                <th style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", border: "1px solid black" }}>Work Item Type</th>
                                <th style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", border: "1px solid black" }}>Title</th>
                                <th style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", border: "1px solid black" }}>State</th>
                                <th style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", border: "1px solid black" }}>Tags</th>
                                <th style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", border: "1px solid black" }}>Area Path</th>
                                <th style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", border: "1px solid black" }}>Iteration Path</th>
                            </tr>
                        </thead>
                        <tbody>
                            {this.state.workItems.map(item => (
                                <tr
                                key={item.id}
                                onClick={() => this.handleRowClick(item)}
                                style={{ cursor: "pointer", backgroundColor: this.state.selectedRow === item.id ? "lightblue" : "white" }}
                                >
                                    <td style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", border: "1px solid black" }}>{item.fields["System.WorkItemType"]}</td>
                                    <td style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", border: "1px solid black" }}>{item.fields["System.Title"]}</td>
                                    <td style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", border: "1px solid black" }}>{item.fields["System.State"]}</td>
                                    <td style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", border: "1px solid black" }}>{item.fields["System.Tags"]}</td>
                                    <td style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", border: "1px solid black" }}>{item.fields["System.AreaPath"]}</td>
                                    <td style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", border: "1px solid black" }}>{item.fields["System.IterationPath"]}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
                <div>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: "10px" }}>
                        {this.state.selectedTags.map(tag => (
                            <div key={tag} style={{ display: "flex", alignItems: "center", background: "#e0e0e0", borderRadius: "15px" }}>
                                <span>{tag}</span>
                                <Button
                                    iconProps={{ iconName: "Cancel" }}
                                    onClick={() => this.removeTag(tag)}
                                    style={{ marginLeft: "5px", minWidth: "20px" }}
                                />
                            </div>
                        ))}
                    </div>
                    <div style={{ display: "flex", marginTop: "10px", gap: "10px" }}>
                        <TextField
                            disabled={!this.state.selectedRow}
                            value={this.state.newTag}
                            onChange={(e, newValue) => this.setState({ newTag: newValue })}
                            placeholder="Add new tag"
                        />
                        <Button disabled={!this.state.newTag} text="Add Tag" onClick={this.addTag} />
                    </div>
                </div>
                <div style={{ display: "flex", justifyContent: "flex-end", gap: "10px" }}>
                    <Button text="Cancel" onClick={() => alert("Cancel not possible")} />
                    <Button text="Save" primary={true} onClick={this.updateTags} />
                </div>
            </Page>
        );
    }

    private async loadProjectContext(): Promise<void> {
        try {
            const client = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
            const context = await client.getProject();

            this.loadWorkItems(context?.name);

            SDK.notifyLoadSucceeded();
        } catch (error) {
            console.error("Failed to load project context: ", error);
        }
    }

    private async loadWorkItems(projectName: string | undefined): Promise<void> {
        if (!projectName) {
            console.error("Project name is not defined");
            return;
        }
        try {
            const client = getClient(WorkItemTrackingRestClient);
            const project = projectName;

            const query = `
                SELECT [System.Id], [System.Title], [System.WorkitemType]
                FROM WorkItems
                WHERE [System.TeamProject] = '${project}'
                AND [System.WorkItemType] IN ('Epic', 'Issue')
                ORDER BY [System.ChangedDate] DESC
            `;

            const result = await client.queryByWiql({ query }, project);
            const workItems = await client.getWorkItems(result.workItems.map(wi => wi.id));

            this.setState({ workItems });

            SDK.notifyLoadSucceeded();
        } catch (error) {
            console.error("Failed to load work items: ", error);
        }
    }
}

showRootComponent(<MenuModal />);