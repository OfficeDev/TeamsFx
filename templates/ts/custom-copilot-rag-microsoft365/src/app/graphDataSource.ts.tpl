import { DataSource, Memory, RenderedPromptSection, Tokenizer } from "@microsoft/teams-ai";
import { TurnContext } from "botbuilder";
import { Client, ResponseType } from "@microsoft/microsoft-graph-client";

/**
 * A data source that searches through Graph API.
 */
export class GraphDataSource implements DataSource {
    /**
     * Name of the data source.
     */
    public readonly name: string;

    /**
     * Name of the external connection.
     */
    public readonly connectionName: string;

    /**
     * Graph client to make requests to Graph API.
     */
    private graphClient: Client;

    /**
     * Creates a new instance of the Graph DataSource instance.
     */
    public constructor(name: string, connectionName: string) {
        this.name = name;
        this.connectionName = connectionName;
    }

    /**
     * Renders the data source as a string of text.
     * @remarks
     * The returned output should be a string of text that will be injected into the prompt at render time.
     * @param context Turn context for the current turn of conversation with the user.
     * @param memory An interface for accessing state values.
     * @param tokenizer Tokenizer to use when rendering the data source.
     * @param maxTokens Maximum number of tokens allowed to be rendered.
     * @returns A promise that resolves to the rendered data source.
     */
    public async renderData(context: TurnContext, memory: Memory, tokenizer: Tokenizer, maxTokens: number): Promise<RenderedPromptSection<string>> {
        const query = memory.getValue("temp.input") as string;
        if(!query) {
            return { output: "", length: 0, tooLong: false };
        }
        if (!this.graphClient) {
            this.graphClient = Client.init({
                authProvider: (done) => {
                    done(null, (memory as any).temp.authTokens["graph"]);
                }
            });
        }
        const graphQuery = query;
        const contentResults = [];
        const response = await this.graphClient.api("/search/query").post({
            requests: [
                {
                    entityTypes: ["externalItem"],
                    contentSources: [
                        `/external/connections/${this.connectionName}`
                    ],
                    query: {
                        queryString: graphQuery,
                    },
                    // This parameter is required only when searching with application permissions
                    // https://learn.microsoft.com/graph/search-concept-searchall
                    // region: "US",
                },
            ],
        });
        for (const value of response?.value ?? []) {
            for (const hitsContainer of value?.hitsContainers ?? []) {
              contentResults.push(...(hitsContainer?.hits ?? []));
            }
        }

        // Add documents until you run out of tokens
        let length = 0,
        output = "";
        for (const result of contentResults) {
            const rawContent = await this
                .downloadExternalContent(result.resource.properties.substrateContentDomainId);
            if (!rawContent) {
                continue;
            }
            let doc = `${rawContent}\n\n`;
            let docLength = tokenizer.encode(doc).length;
            const remainingTokens = maxTokens - (length + docLength);
            if (remainingTokens <= 0) {
                break;
            }

            // Append do to output
            output += doc;
            length += docLength;
        }
        return { output: this.formatDocument(output), length: output.length, tooLong: false };
    }

    /**
     * Formats the result string 
     * @param result 
     * @returns 
     */
    private formatDocument(result: string): string {
        return `<context>${result}</context>`;
    }

    /**
     * Download external item content
     * @param externalItemFullId Full ID of the external item 
     * @returns External item content
     */
    private async downloadExternalContent(externalItemFullId: string): Promise<string> {
        const externalItemId = externalItemFullId.split(',')[1];
        const externalItem = await this.graphClient
            .api(`/external/connections/${this.connectionName}/items/${externalItemId}`)
            .get();
        return externalItem.content.value;
    }
}