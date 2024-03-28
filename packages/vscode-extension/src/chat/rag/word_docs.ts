// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
export const wordJsApiDocs = [
  {
    objeName: "Word.Body",
    apiList: [
      {
        name: "Word.Body.getComments",
        description: "Get all the comments in the document body.",
        kind: "Method",
        signature: "Word.Body.getComments(): Word.CommentCollection",
        examples: [
          'const comments = context.document.body.getComments(); \n   comments.load("content, items, replies"); \n  await context.sync();',
        ],
      },
      {
        name: "Word.Body.getHtml",
        description: "Gets an HTML representation of the body object. ",
        kind: "Method",
        signature: "Word.Body.getHtml(): OfficeExtension.ClientResult<string>",
        examples: [],
      },
    ],
  },
  {
    objeName: "Word.Range",
    apiList: [
      {
        name: "Word.Range.getComments",
        description: "Get all the comments in the range or selection.",
        kind: "Method",
        signature: "Word.Range.getComments(): Word.CommentCollection",
        examples: [
          'const comments = context.document.getSelection().getComments(); \n   comments.load("content, items, replies"); \n  await context.sync();',
        ],
      },
      {
        name: "Word.Range.getHtml",
        description: "Gets an HTML representation of the range object or current selection. ",
        kind: "Method",
        signature: "Word.Range.getHtml(): OfficeExtension.ClientResult<string>",
        examples: [],
      },
    ],
  },
  {
    objeName: "Word.Paragraph",
    apiList: [
      {
        name: "Word.Paragraph.getComments",
        description: "Get all the comments in the paragraph.",
        kind: "Method",
        signature: "Word.Paragraph.getComments(): Word.CommentCollection",
        examples: [
          'const comments = context.document.paragraphs.getFirst().getComments(); \n   comments.load("content, items, replies"); \n  await context.sync();',
        ],
      },
      {
        name: "Word.Paragraph.getHtml",
        description: "Gets an HTML representation of the paragraph.",
        kind: "Method",
        signature: "Word.Paragraph.getHtml(): OfficeExtension.ClientResult<string>",
        examples: [],
      },
    ],
  },
  {
    objeName: "Word.Comment",
    apiList: [
      {
        name: "Word.Comment.authorEmail",
        description: "Get the email of the comment's author",
        kind: "Property",
        signature: "Word.Comment.authorEmail: string",
        examples: [],
      },
      {
        name: "Word.Comment.authorName",
        description: "Gets the name of the comment's author.",
        kind: "Property",
        signature: "Word.Comment.authorName: string",
        examples: [],
      },
      {
        name: "Word.Comment.content",
        description: "get or set the comment's content as plain text.",
        kind: "Property",
        signature: "Word.Comment.content",
        examples: [
          "const comment = context.document.getSelection().getComments().getFirst();\n comment.content = text;\n",
        ],
      },
      {
        name: "Word.Comment.creationDate",
        description: "Gets the creation date of the comment",
        kind: "Property",
        signature: "Word.Comment.creationDate: string",
        examples: [
          'const comment = context.document.getSelection().getComments().getFirst();\n comment.load("creationDate");\n',
        ],
      },
      {
        name: "Word.Comment.replies",
        description: "Gets the collection of reply objects associated with the comment.",
        kind: "Property",
        signature: "Word.Comment.replies: Word.CommentReplyCollection",
        examples: [],
      },
      {
        name: "Word.Comment.resolved",
        description:
          "Specifies the comment thread's status. Setting to true resolves the comment thread. Getting a value of true means that the comment thread is resolved.",
        kind: "Property",
        signature: "Word.Comment.resolved: boolean",
        examples: [
          "const comment = context.document.getSelection().getComments().getFirst();\n comment.resolved = true;\n",
        ],
      },
      {
        name: "Word.Comment.delete",
        description: "Deletes the comment and its replies.",
        kind: "Method",
        signature: "Word.Comment.delete: void",
        examples: [
          "const comment = context.document.getSelection().getComments().getFirst();\n comment.delete();\n",
        ],
      },
      {
        name: "Word.Comment.reply",
        description: "Reply the comment and its replies.",
        kind: "Method",
        signature: "Word.Comment.reply(replyText: string): Word.CommentReply",
        examples: [
          ' const comments = context.document.getSelection().getComments();\n  comments.load("items");\n  await context.sync();\n  const firstActiveComment = comments.items.find((item) => item.resolved !== true);\n   if (firstActiveComment) { \n   const reply = firstActiveComment.reply(text); \n    console.log("Reply added"); }',
        ],
      },
      {
        name: "Word.Comment.getRange",
        description: "Gets the range in the main document where the comment is on.",
        kind: "Method",
        signature: "Word.Comment.getRange(): Word.Range",
        examples: [
          " const range = context.document.getSelection().getComments().getFirst().getRange(); \n range.load();\n  await context.sync();",
        ],
      },
    ],
  },
];
