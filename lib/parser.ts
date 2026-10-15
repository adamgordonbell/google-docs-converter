const util = require('./utilities');
const writers = require('./writers');
import {Writer} from './writers';

/**
 * Helper to get all tabs recursively (including nested child tabs)
 */
function getAllTabs(tabs: object[], result: object[] = []): object[] {
  for (const tab of tabs || []) {
    result.push(tab);
    if (tab['childTabs']) {
      getAllTabs(tab['childTabs'], result);
    }
  }
  return result;
}

/**
 * Function to parse Google Docs API Document
 * @param {object} document JSON object representing document
 * @param {Writer} writer Converter to appropriate text format
 * @return {string}
 */
function parseDocument(
    document: object,
    writer: Writer = new writers.MarkdownWriter(),
): string {
  // Handle multi-tab documents (when includeTabsContent is true)
  if (document['tabs'] && document['tabs'].length > 0) {
    const allTabs = getAllTabs(document['tabs']);
    let allParsed: string[] = [];

    for (const tab of allTabs) {
      const docTab = tab['documentTab'];
      if (!docTab || !docTab['body']) continue;

      const body: object = docTab['body'];
      const content: object[] = body['content'];
      const lists: object = docTab['lists'] || {};
      const listTracker: object = {};

      // Add tab title as header if multiple tabs
      if (allTabs.length > 1 && tab['tabProperties'] && tab['tabProperties']['title']) {
        allParsed.push(writer.addHeading(tab['tabProperties']['title'], 1));
      }

      const parsed = parseContent(content, writer, lists, listTracker);
      allParsed = allParsed.concat(parsed);
    }

    return writer.finalize(allParsed);
  }

  // Fallback for single-tab or legacy documents
  const body: object = document['body'];
  const content: object[] = body['content'];
  const lists: object = document['lists'];
  const listTracker: object = {};

  const parsed: string[] = parseContent(content, writer, lists, listTracker);
  const convertedFinal: string = writer.finalize(parsed);
  return convertedFinal;
}

exports.parseDocument = parseDocument;

/**
 * Process content array items in order, handling both paragraphs and tables
 */
function parseContent(
    content: object[],
    writer: Writer,
    lists: object,
    listTracker: object,
): string[] {
  const result: string[] = [];
  for (const item of content) {
    if (item.hasOwnProperty('paragraph')) {
      result.push(parseParagraph(item['paragraph'], writer, lists, listTracker));
    } else if (item.hasOwnProperty('table')) {
      result.push(parseTable(item['table'], writer, lists, listTracker));
    }
  }
  return result;
}

/**
 * Parse a Google Docs table into a markdown table
 */
function parseTable(
    table: object,
    writer: Writer,
    lists: object,
    listTracker: object,
): string {
  const tableRows: object[] = table['tableRows'] || [];
  const rows: string[][] = [];

  for (const row of tableRows) {
    const cells: string[] = [];
    const tableCells: object[] = row['tableCells'] || [];
    for (const cell of tableCells) {
      const cellContent: object[] = cell['content'] || [];
      const cellParts: string[] = [];
      for (const item of cellContent) {
        if (item.hasOwnProperty('paragraph')) {
          const text = parseParagraph(item['paragraph'], writer, lists, listTracker);
          const trimmed = text.replace(/\n/g, ' ').trim();
          if (trimmed) cellParts.push(trimmed);
        }
      }
      cells.push(cellParts.join(' '));
    }
    rows.push(cells);
  }

  if (rows.length === 0) return '';

  // Build markdown table
  const lines: string[] = [];
  // Header row
  lines.push('| ' + rows[0].join(' | ') + ' |');
  // Separator
  lines.push('| ' + rows[0].map(() => '---').join(' | ') + ' |');
  // Data rows
  for (let i = 1; i < rows.length; i++) {
    lines.push('| ' + rows[i].join(' | ') + ' |');
  }
  return '\n' + lines.join('\n') + '\n';
}

/** Function to parse paragraph element and convert to Markdown
 * @param  {object} paragraph From 'paragraph' field
 * @param  {Writer} writer class to handle conversions, default is markdown
 * @param  {object} lists object to provide id's and details for lists
 * @param  {object} listTracker object to track place in a list
 * @return {string} Markdown as string
 */
function parseParagraph(
    paragraph: object,
    writer: Writer = new writers.MarkdownWriter(),
    lists?: object,
    listTracker?: object,
): string {
  const elements : object[] = paragraph['elements'];
  /** Function to reduce elements/textRuns into single string
   * @param {string} acc
   * @param {object} element
   * @return {string}
  */
  function elementsReducer(acc: string, element: object): string {
    return acc + parseElement(element, writer);
  }
  let content = elements.reduce(elementsReducer, '');

  // Handle Headers
  const paragraphStyle = paragraph['paragraphStyle'];
  const namedStyle: string = paragraphStyle['namedStyleType'];

  if (namedStyle.startsWith('HEADING')) {
    const levelString: string = namedStyle.substring(namedStyle.length - 1);
    const level: number = parseInt(levelString);
    content = writer.addHeading(content, level);
  }

  // Handle Lists
  if (paragraph.hasOwnProperty('bullet')) { // Means it's a list-item
    const bullet = paragraph['bullet'];
    const listId = bullet['listId'];
    const nestingLevel =
      bullet['nestingLevel'] === undefined ? 0 : bullet['nestingLevel'];
    const padding = '    '.repeat(nestingLevel);
    const listDetails =
      lists[listId]['listProperties']['nestingLevels'][nestingLevel];
    if (listDetails.hasOwnProperty('glyphType')) { // means ordered list
      if (listTracker.hasOwnProperty(listId) &&
        listTracker[listId].hasOwnProperty(nestingLevel)) {
        listTracker[listId][nestingLevel] += 1;
      } else {
        /** If no listId entry, create one... */
        if (!listTracker.hasOwnProperty(listId)) {
          listTracker[listId] = {};
        }
        listTracker[listId][nestingLevel] = 1;
      }
      content = listTracker[listId][nestingLevel].toString() + '. ' + content;
    } else { // means unordered list
      content = '- ' + content;
    }
    // Apply padding using nestingLevel
    content = padding + content;
    // Put newline at end of line, rather than within style text
    content = content.replace('\n', '') + '\n';
  }
  return content;
}

exports.parseParagraph = parseParagraph;

/** Function to parse individual element into appropriate format
 * @param {object} element
 * @param {Writer} writer to use for conversion (markdown, org, etc)
 * @return {string}
 */
function parseElement(element: object, writer: Writer): string {
  // Currently, parseElement only supports the textRun form of a 
  // ParagraphElement 
  // SEE: https://developers.google.com/docs/api/reference/rest/v1/documents#ParagraphElement
  if (!element.hasOwnProperty('textRun')) { // means NOT regular paragraph
    return ""
  }
  let content = element['textRun']['content'];
  const textStyle = element['textRun']['textStyle'];
  if (util.isEmpty(textStyle)) {
    return content;
  }

  /* Note, the order of these conditionals matters.
   * HTML tags on outside of markdown. */
  if (textStyle.underline && !textStyle.link) { // Ignore underlining if a link
    content = writer.underline(content);
  }
  if (textStyle.strikethrough) {
    content = writer.strikethrough(content);
  }
  if (textStyle.italic) {
    content = writer.italicize(content);
  }
  if (textStyle.bold) {
    content = writer.bold(content);
  }
  if (textStyle.link) {
    if (textStyle.link.url) {
      content = writer.addLink(content, textStyle.link.url);
    }
  }
  return content;
}
