"""
Critical Path Method (CPM) calculation engine.
Performs Forward Pass, Backward Pass, Float Calculation, and Critical Path extraction.

The forward pass is calendar aware: an activity's EFFORT days (productivity
formula) are converted to CALENDAR days using the start date the forward pass
gave it, so weekly rest days, public holidays and the monsoon slowdown all push
the finish date out. Duration is therefore fixed once scheduled, and the
backward pass uses exactly that duration.

The critical path is traced backwards through real predecessor edges rather than
by sorting every zero-float activity, so the reported chain is an actual
connected path through the network.
"""
from typing import List, Dict, Optional, Tuple
from datetime import timedelta, date
from .models import RoadActivity, ProjectInput
from .project_calendar import ProjectCalendar


class CPMEngine:
    def __init__(self, activities: List[RoadActivity], project_input: ProjectInput,
                 calendar: Optional[ProjectCalendar] = None):
        self.activities = activities
        self.project_input = project_input
        self.calendar = calendar
        self.act_dict: Dict[str, RoadActivity] = {a.activity_id: a for a in activities}
        self.successors: Dict[str, List[str]] = {a.activity_id: [] for a in activities}

        for act in activities:
            for pred in act.predecessors:
                if pred in self.successors:
                    self.successors[pred].append(act.activity_id)

    # ------------------------------------------------------------------ dates
    def _calendar_duration(self, act: RoadActivity, start_date: date) -> int:
        """Effort days -> calendar days, honouring the project calendar."""
        effort = float(act.effort_days if act.effort_days > 0 else act.duration_days)
        if effort <= 0:
            return 0

        if self.calendar is None:
            # Without a calendar, effort days are used directly (legacy mode).
            return max(1, int(round(effort)))
        return self.calendar.calendar_days_needed(start_date, effort)

    def run_cpm(self) -> Tuple[int, List[str]]:
        """
        Executes Forward and Backward Pass.
        Returns (project_duration_days, critical_path_sequence).
        """
        start_dt = self.project_input.start_date

        # 1. Topological Sorting (Kahn's algorithm)
        in_degree = {a.activity_id: len(a.predecessors) for a in self.activities}
        queue = [a.activity_id for a in self.activities if in_degree[a.activity_id] == 0]
        topo_order = []

        while queue:
            curr = queue.pop(0)
            topo_order.append(curr)
            for nxt in self.successors.get(curr, []):
                in_degree[nxt] -= 1
                if in_degree[nxt] == 0:
                    queue.append(nxt)

        # 2. Forward Pass: ES/EF in day-ordinals, duration fixed by the calendar.
        for act_id in topo_order:
            act = self.act_dict[act_id]
            if not act.predecessors:
                act.es = 0
            else:
                act.es = max(
                    max(0, self.act_dict[p].ef + act.predecessor_lags.get(p, 0))
                    for p in act.predecessors
                )
            act.duration_days = self._calendar_duration(
                act, start_dt + timedelta(days=act.es))
            act.ef = act.es + act.duration_days

        project_duration = max((act.ef for act in self.activities), default=0)

        # 3. Backward Pass: LF, LS, Total Float, Free Float.
        #    Float is left UNCLAMPED: a negative float means the activity cannot
        #    finish in time and must be reported, not silently zeroed.
        for act_id in reversed(topo_order):
            act = self.act_dict[act_id]
            succs = self.successors.get(act_id, [])
            if not succs:
                act.lf = project_duration
                act.free_float = project_duration - act.ef
            else:
                act.lf = min(
                    self.act_dict[s].ls - self.act_dict[s].predecessor_lags.get(act_id, 0)
                    for s in succs
                )
                act.free_float = min(
                    (self.act_dict[s].es - self.act_dict[s].predecessor_lags.get(act_id, 0))
                    for s in succs
                ) - act.ef

            act.ls = act.lf - act.duration_days
            act.total_float = act.ls - act.es
            # A zero-duration, zero-quantity activity is OUT OF SCOPE (e.g. an
            # undeclared HDD/river crossing or an absent treatment plant). It
            # must not be reported as a critical activity: it contributes no
            # work and no calendar, so calling it critical would put a phantom
            # node on the reported critical path.
            act.is_critical = (act.total_float == 0 and act.duration_days > 0)

        # 4. Map Dates
        for act in self.activities:
            act.planned_start = start_dt + timedelta(days=act.es)
            act.planned_finish = start_dt + timedelta(days=act.ef)

        # 5. Extract the critical path as a real chain of predecessor edges.
        critical_path_ids = self._trace_critical_chain(project_duration)

        return project_duration, critical_path_ids

    def _trace_critical_chain(self, project_duration: int) -> List[str]:
        """
        Walk backwards from the activity that determines the project duration,
        following only real predecessor edges where that predecessor is the one
        that actually drives the successor's early start.

        Returns the chain in start-to-finish order. Other zero-float activities
        that are not on this chain belong to PARALLEL critical paths and are
        exposed through `parallel_critical_ids`.
        """
        critical = [a for a in self.activities if a.is_critical]
        self.parallel_critical_ids: List[str] = []
        if not critical:
            return []

        # A zero-duration activity is a PASS-THROUGH node (typically a
        # material/size branch that is out of scope, or a declared crossing
        # that was not configured). It is not a work-critical activity, but it
        # is still a legitimate connector when the driving chain passes
        # through it. Traverse those nodes, but keep them out of the reported
        # work chain.
        def drives(a: RoadActivity) -> bool:
            return bool(a.is_critical or (a.duration_days == 0 and a.total_float == 0))

        terminal = [a for a in critical if a.ef == project_duration]
        if not terminal:
            terminal = [max(critical, key=lambda a: a.ef)]
        # Longest chain first: the terminal critical activity with the latest ES.
        current = max(terminal, key=lambda a: (a.es, a.ef))

        chain = [current]
        seen = {current.activity_id}
        while current.predecessors:
            drivers = []
            for pid in current.predecessors:
                pred = self.act_dict.get(pid)
                if pred is None or not drives(pred) or pid in seen:
                    continue
                lag = current.predecessor_lags.get(pid, 0)
                if pred.ef + lag == current.es:
                    drivers.append(pred)
            if not drivers:
                break
            # Prefer a real work activity; fall back to a pass-through node and
            # keep walking until the chain re-joins real work.
            nxt = max(drivers, key=lambda a: (a.duration_days > 0, a.ef, a.es))
            seen.add(nxt.activity_id)
            if nxt.duration_days > 0:
                chain.append(nxt)
            current = nxt

        chain.reverse()
        self.parallel_critical_ids = [
            a.activity_id for a in critical if a.activity_id not in seen
        ]
        return [a.activity_id for a in chain]
