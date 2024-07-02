const { Client, ResponseType } = require("@microsoft/microsoft-graph-client");

/**
 * A data source that searches through Graph API.
 */
class GraphDataSource {
    /**
     * Creates a new instance of the Graph DataSource instance.
     */
    constructor(name, connectionName) {
        this.name = name;
        this.connectionName = connectionName;
    }

    /**
     * Renders the data source as a string of text.
     */
    async renderData(context, memory, tokenizer, maxTokens) {
        const query = memory.getValue("temp.input");
        if(!query) {
            return { output: "", length: 0, tooLong: false };
        }
        if (!this.graphClient) {
            this.graphClient = Client.init({
                authProvider: (done) => {
                    done(null, memory.temp.authTokens["graph"]);
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
     */
    formatDocument(result) {
        return `<context>${result}</context>`;
    }

    async downloadExternalContent(externalItemFullId) {
        const externalItemId = externalItemFullId.split(',')[1];
        const externalItem = await this.graphClient
            .api(`/external/connections/${this.connectionName}/items/${externalItemId}`)
            .get();
        return externalItem.content.value;
    }
}

module.exports = {
  GraphDataSource,
};