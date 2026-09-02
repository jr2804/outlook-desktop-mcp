"""Task operations on Windows via COM."""

from __future__ import annotations

from collections.abc import Callable
from typing import Any, cast

from outlook_desktop_mcp.backends.win._types import Namespace, TaskItem
from outlook_desktop_mcp.backends.win.formatting import (
    format_task_full,
    format_task_summary,
)
from outlook_desktop_mcp.backends.win.helpers import (
    _OL_CLASS_TASK,
    _require_class,
    _require_store,
)
from outlook_desktop_mcp.models import (
    ItemStatusResult,
    TaskCreatedResult,
    TaskFull,
    TaskSummary,
)
from outlook_desktop_mcp.tools._folder_constants import (
    OL_FOLDER_TASKS,
    OL_TASK_COMPLETE,
    OL_TASK_ITEM,
)


async def list_tasks(
    bridge: Any,
    include_completed: bool,
    count: int,
    account: str,
) -> list[TaskSummary]:
    def _list(outlook: Any, namespace: Namespace) -> list[TaskSummary]:
        effective_count = min(max(1, count), 200)
        store = _require_store(namespace, account)
        folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
        items = folder.Items
        items.Sort("[DueDate]")

        if not include_completed:
            items = items.Restrict("[Complete] = False")

        results = []
        limit = min(effective_count, items.Count)
        for i in range(limit):
            try:
                results.append(TaskSummary.model_validate(format_task_summary(items.Item(i + 1))))
            except Exception:  # noqa: S112
                continue
        return results

    return await bridge.call(_list)


async def get_task(bridge: Any, get_item: Callable, entry_id: str, account: str) -> TaskFull:
    def _get(outlook: Any, namespace: Namespace) -> TaskFull:
        item = get_item(namespace, entry_id, account)
        return TaskFull.model_validate(format_task_full(item))

    return await bridge.call(_get)


async def create_task(
    bridge: Any,
    subject: str,
    body: str,
    due_date: str,
    importance: str,
    reminder_minutes: int,
    account: str,
) -> TaskCreatedResult:
    def _create(outlook: Any, namespace: Namespace) -> TaskCreatedResult:
        task: TaskItem = outlook.CreateItem(OL_TASK_ITEM)
        if account:
            store = _require_store(namespace, account)
            tasks_folder = store.GetDefaultFolder(OL_FOLDER_TASKS)
            task = cast(TaskItem, task.Move(tasks_folder))
        task.Subject = subject
        if body:
            task.Body = body
        if due_date:
            task.DueDate = due_date
        imp_map = {"low": 0, "normal": 1, "high": 2}
        task.Importance = imp_map.get(importance.lower(), 1)
        if reminder_minutes > 0:
            task.ReminderSet = True
            task.ReminderMinutesBeforeStart = reminder_minutes
        else:
            task.ReminderSet = False
        task.Save()
        return TaskCreatedResult(
            subject=task.Subject,
            entry_id=task.EntryID,
            due_date=str(task.DueDate) if due_date else None,
        )

    return await bridge.call(_create)


async def complete_task(bridge: Any, get_item: Callable, entry_id: str, account: str) -> ItemStatusResult:
    def _complete(outlook: Any, namespace: Namespace) -> ItemStatusResult:
        item = cast(TaskItem, get_item(namespace, entry_id, account))
        _require_class(item, _OL_CLASS_TASK, "task item")
        item.Status = OL_TASK_COMPLETE
        item.PercentComplete = 100
        item.Save()
        return ItemStatusResult(status="completed", subject=item.Subject, entry_id=entry_id)

    return await bridge.call(_complete)


async def delete_task(bridge: Any, get_item: Callable, entry_id: str, account: str) -> ItemStatusResult:
    def _delete(outlook: Any, namespace: Namespace) -> ItemStatusResult:
        item = cast(TaskItem, get_item(namespace, entry_id, account))
        _require_class(item, _OL_CLASS_TASK, "task item")
        subject = item.Subject
        item.Delete()
        return ItemStatusResult(status="deleted", subject=subject, entry_id=entry_id)

    return await bridge.call(_delete)
