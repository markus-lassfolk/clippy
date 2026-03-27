/**
 * Convert basic markdown to HTML for email.
 * Supports: bold, italic, links, unordered lists, ordered lists, line breaks.
 */
export function markdownToHtml(text: string): string {
  let html = text;

  // Escape HTML entities first
  html = html.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  // Bold: **text** or __text__
  html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/__(.+?)__/g, '<strong>$1</strong>');

  // Italic: *text* or _text_ (but not inside words)
  html = html.replace(/(?<!\w)\*([^*]+?)\*(?!\w)/g, '<em>$1</em>');
  html = html.replace(/(?<!\w)_([^_]+?)_(?!\w)/g, '<em>$1</em>');

  // Links: [text](url)
  html = html.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>');

  // Process lists - need to handle line by line
  const lines = html.split('\n');
  const result: string[] = [];
  let inUnorderedList = false;
  let inOrderedList = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const unorderedMatch = line.match(/^[\s]*[-*]\s+(.+)$/);
    const orderedMatch = line.match(/^[\s]*\d+\.\s+(.+)$/);

    if (unorderedMatch) {
      if (inOrderedList) {
        result.push('</ol>');
        inOrderedList = false;
      }
      if (!inUnorderedList) {
        result.push('<ul>');
        inUnorderedList = true;
      }
      result.push(`<li>${unorderedMatch[1]}</li>`);
    } else if (orderedMatch) {
      if (inUnorderedList) {
        result.push('</ul>');
        inUnorderedList = false;
      }
      if (!inOrderedList) {
        result.push('<ol>');
        inOrderedList = true;
      }
      result.push(`<li>${orderedMatch[1]}</li>`);
    } else {
      // Close any open lists
      if (inUnorderedList) {
        result.push('</ul>');
        inUnorderedList = false;
      }
      if (inOrderedList) {
        result.push('</ol>');
        inOrderedList = false;
      }
      result.push(line);
    }
  }

  // Close any remaining open lists
  if (inUnorderedList) {
    result.push('</ul>');
  }
  if (inOrderedList) {
    result.push('</ol>');
  }

  html = result.join('\n');

  // Convert line breaks to <br> (but not inside lists)
  // Split by list tags, process non-list parts
  html = html.replace(/\n(?!<\/?[uo]l>|<\/?li>)/g, '<br>\n');

  // Wrap in basic HTML structure for email
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; line-height: 1.5; }
  a { color: #0066cc; }
  ul, ol { margin: 10px 0; padding-left: 20px; }
  li { margin: 5px 0; }
</style>
</head>
<body>
${html}
</body>
</html>`;
}

/**
 * Check if text contains markdown formatting.
 */
export function hasMarkdown(text: string): boolean {
  // Check for common markdown patterns
  return /\*\*.+?\*\*|__.+?__|(?<!\w)\*[^*]+?\*(?!\w)|(?<!\w)_[^_]+?_(?!\w)|\[.+?\]\(.+?\)|^[\s]*[-*]\s+|^[\s]*\d+\.\s+/m.test(
    text
  );
}
