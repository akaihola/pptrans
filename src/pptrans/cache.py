"""Cache management for translations."""

import hashlib
import json
import os

import click


def load_cache(cache_file_path) -> dict:
    """Loads the translation cache from a JSON file."""
    if os.path.exists(cache_file_path):
        try:
            with open(cache_file_path, encoding="utf-8") as f:
                return json.load(f)
        except (OSError, json.JSONDecodeError) as e:
            click.echo(
                f"Warning: Could not load cache file {cache_file_path}. Error: {e}. Starting with an empty cache.",
                err=True,
            )
    return {}


def save_cache(cache_data, cache_file_path) -> None:
    """Saves the translation cache to a JSON file."""
    try:
        with open(cache_file_path, "w", encoding="utf-8") as f:
            json.dump(cache_data, f, indent=4, ensure_ascii=False)
    except OSError as e:
        click.echo(
            f"Warning: Could not save cache file {cache_file_path}. Error: {e}",
            err=True,
        )


def generate_page_hash(texts_on_page: list[str]) -> str:
    """Generates a SHA256 hash for a list of text strings from a page."""
    concatenated_texts = "|".join(texts_on_page)  # Use a delimiter
    return hashlib.sha256(concatenated_texts.encode("utf-8")).hexdigest()
