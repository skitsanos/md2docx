"""
Markdown parser using mistune to generate AST.

This module parses Markdown content into an Abstract Syntax Tree (AST)
that can be processed by the Word document generator.
"""

from typing import Any
import mistune


class MarkdownParser:
    """
    Parse Markdown content into an Abstract Syntax Tree (AST).

    Uses mistune's built-in AST renderer to create a tree structure
    that represents the Markdown document.
    """

    def __init__(self):
        """Initialize the parser with mistune's AST renderer and plugins."""
        self._markdown = mistune.create_markdown(
            renderer="ast",
            plugins=["table", "strikethrough", "task_lists", "def_list"]
        )

    def parse(self, content: str) -> list[dict[str, Any]]:
        """
        Parse Markdown content into an AST.

        Args:
            content: Markdown text content to parse.

        Returns:
            A list of AST nodes representing the document structure.
            Each node is a dictionary with at least a 'type' key.

        Example:
            >>> parser = MarkdownParser()
            >>> ast = parser.parse("# Hello\\n\\nThis is a **test**.")
            >>> ast[0]['type']
            'heading'
        """
        return self._markdown(content)

    def parse_file(self, file_path: str) -> list[dict[str, Any]]:
        """
        Parse a Markdown file into an AST.

        Args:
            file_path: Path to the Markdown file.

        Returns:
            A list of AST nodes representing the document structure.

        Raises:
            FileNotFoundError: If the file does not exist.
            IOError: If the file cannot be read.
        """
        with open(file_path, "r", encoding="utf-8") as f:
            content = f.read()
        return self.parse(content)

    @staticmethod
    def get_node_type(node: dict[str, Any]) -> str:
        """
        Get the type of an AST node.

        Args:
            node: An AST node dictionary.

        Returns:
            The type string of the node.
        """
        return node.get("type", "unknown")

    @staticmethod
    def get_node_children(node: dict[str, Any]) -> list[dict[str, Any]]:
        """
        Get the children of an AST node.

        Args:
            node: An AST node dictionary.

        Returns:
            A list of child nodes, or an empty list if none.
        """
        return node.get("children", [])

    @staticmethod
    def get_node_text(node: dict[str, Any]) -> str:
        """
        Extract text content from an AST node recursively.

        Args:
            node: An AST node dictionary.

        Returns:
            The concatenated text content of the node and its children.
        """
        if "raw" in node:
            return node["raw"]

        if "children" in node:
            return "".join(
                MarkdownParser.get_node_text(child)
                for child in node["children"]
            )

        return ""

    @staticmethod
    def pretty_print_ast(ast: list[dict[str, Any]], indent: int = 0) -> str:
        """
        Create a pretty-printed string representation of the AST.

        Args:
            ast: The AST to print.
            indent: Current indentation level.

        Returns:
            A formatted string representation of the AST.
        """
        lines = []
        indent_str = "  " * indent

        for node in ast:
            node_type = node.get("type", "unknown")
            lines.append(f"{indent_str}{node_type}")

            # Add relevant attributes
            for key, value in node.items():
                if key not in ("type", "children"):
                    if isinstance(value, str) and len(value) > 50:
                        value = value[:50] + "..."
                    lines.append(f"{indent_str}  {key}: {value}")

            # Recursively print children
            if "children" in node:
                lines.append(
                    MarkdownParser.pretty_print_ast(node["children"], indent + 1)
                )

        return "\n".join(lines)
