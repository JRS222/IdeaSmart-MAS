import markdown
import re

def convert_markdown_to_html(markdown_file, html_template_file, output_file):
    # Read the markdown content
    with open(markdown_file, 'r', encoding='utf-8') as f:
        md_content = f.read()

    # Convert Mermaid blocks to HTML
    md_content = re.sub(
        r'```mermaid\n(.*?)\n```',
        r'<div class="mermaid-container">\n<div class="mermaid">\n\1\n</div>\n</div>',
        md_content,
        flags=re.DOTALL
    )

    # Convert image references from ![[image.png]] to proper HTML img tags
    md_content = re.sub(
        r'!\[\[(.*?)\]\]',
        r'<div class="image-container">\n<img src="images/\1" alt="\1">\n</div>',
        md_content
    )

    # Convert markdown to HTML
    html_content = markdown.markdown(
        md_content,
        extensions=['extra', 'toc']
    )

    # Read the HTML template
    with open(html_template_file, 'r', encoding='utf-8') as f:
        template = f.read()

    # Replace placeholder with converted content
    final_html = template.replace('{{CONTENT}}', html_content)

    # Write the final HTML
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(final_html)

# Create HTML template
template_html = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Parts Management Documentation</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/mermaid/10.6.1/mermaid.min.js"></script>
    <link rel="stylesheet" href="styles.css">
</head>
<body>
    <div class="container">
        <nav id="sidebar">
            <div class="nav-header">Table of Contents</div>
            <ul>
                <li><a href="#introduction" class="nav-link">Introduction</a></li>
                <li><a href="#parts-books" class="nav-link">Parts Books</a></li>
                <li><a href="#call-logs" class="nav-link">Call Logs</a></li>
                <li><a href="#labor-log" class="nav-link">Labor Log</a></li>
                <li><a href="#actions" class="nav-link">Actions</a></li>
                <li><a href="#search" class="nav-link">Search</a></li>
                <li><a href="#workflow" class="nav-link">Workflow</a></li>
                <li><a href="#powershell" class="nav-link">PowerShell</a></li>
            </ul>
        </nav>

        <main id="content">
            {{CONTENT}}
        </main>
    </div>
    <script src="script.js"></script>
</body>
</html>
"""

# Save template
with open('template.html', 'w', encoding='utf-8') as f:
    f.write(template_html)

# Install required package if not already installed
# pip install markdown

# Usage
if __name__ == "__main__":
    convert_markdown_to_html(
        'documentation.md',  # Your markdown file
        'template.html',     # Template file
        'index.html'         # Output file
    )