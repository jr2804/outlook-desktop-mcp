"""Rule operations on Windows via COM."""

from __future__ import annotations

import logging
from typing import Any

from outlook_desktop_mcp.backends.base import BackendError
from outlook_desktop_mcp.backends.win._types import Namespace
from outlook_desktop_mcp.backends.win.helpers import _require_store
from outlook_desktop_mcp.models import RuleInfo, RuleToggledResult

logger = logging.getLogger(__name__)


async def list_rules(bridge: Any, account: str) -> list[RuleInfo]:
    def _list(outlook: Any, namespace: Namespace) -> list[RuleInfo]:
        store = _require_store(namespace, account)
        rules = store.GetRules()
        return [RuleInfo(index=i + 1, name=rule.Name, enabled=bool(rule.Enabled)) for i in range(rules.Count) for rule in [rules.Item(i + 1)]]

    return await bridge.call(_list)


async def toggle_rule(bridge: Any, rule_name: str, enabled: bool, account: str) -> RuleToggledResult:
    def _toggle(outlook: Any, namespace: Namespace) -> RuleToggledResult:
        store = _require_store(namespace, account)
        rules = store.GetRules()
        for i in range(rules.Count):
            rule = rules.Item(i + 1)
            if rule.Name == rule_name:
                logger.warning("toggle_rule: setting rule '%s' enabled=%s", rule_name, enabled)
                rule.Enabled = enabled
                rules.Save()
                return RuleToggledResult(status="enabled" if enabled else "disabled", rule=rule_name)
        raise BackendError(f"Rule '{rule_name}' not found. Use list_rules to see available rules.")

    return await bridge.call(_toggle)
