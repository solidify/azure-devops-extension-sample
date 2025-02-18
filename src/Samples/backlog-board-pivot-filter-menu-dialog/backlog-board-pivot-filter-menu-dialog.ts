import "es6-promise/auto";
import * as SDK from "azure-devops-extension-sdk";
import { CommonServiceIds, IHostPageLayoutService } from "azure-devops-extension-api";

SDK.register("backlog-board-pivot-filter-menu-dialog", () => {
    return {
        execute: async (context: any) => {
            const dialogService = await SDK.getService<IHostPageLayoutService>(CommonServiceIds.HostPageLayoutService);
            dialogService.openCustomDialog<boolean | undefined>(SDK.getExtensionContext().id + ".menu-modal", {
                title: "Custom board pivot menu dialog",
            });
        }
    }
});

SDK.init();