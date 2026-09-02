"""Category operations on Windows via COM."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

from outlook_desktop_mcp.backends.win._types import Namespace
from outlook_desktop_mcp.models import CategoriesSetResult, CategoryInfo


async def list_categories(bridge: Any, account: str) -> list[CategoryInfo]:
    def _list(outlook: Any, namespace: Namespace) -> list[CategoryInfo]:
        # Categories are profile-wide, not per-store; account accepted for consistency
        return [CategoryInfo(name=cat.Name, color=cat.Color) for i in range(namespace.Categories.Count) for cat in [namespace.Categories.Item(i + 1)]]

    return await bridge.call(_list)


async def set_category(
    bridge: Any,
    get_item: Callable,
    entry_id: str,
    categories: str,
    account: str,
) -> CategoriesSetResult:
    def _set(outlook: Any, namespace: Namespace) -> CategoriesSetResult:
        item = get_item(namespace, entry_id, account)
        item.Categories = categories
        item.Save()
        return CategoriesSetResult(subject=item.Subject, categories=item.Categories or "")

    return await bridge.call(_set)
