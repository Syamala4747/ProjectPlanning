"""
Hierarchical Work Breakdown Structure (WBS) generator and roll-up engine.
Dynamically builds a 3-level WBS tree with cost, duration, and weight roll-ups.
"""
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any
from datetime import date
from .models import RoadActivity

@dataclass
class WBSNode:
    wbs_code: str
    parent_wbs: str
    level: int  # 1: Project, 2: Work Package, 3: Activity
    name: str
    description: str
    activity_id: str = ""
    is_work_package: bool = False
    
    # Roll-up metrics
    quantity: Optional[float] = None
    unit: str = ""
    baseline_cost: float = 0.0
    optimized_cost: float = 0.0
    baseline_cost_crore: float = 0.0
    optimized_cost_crore: float = 0.0
    duration_days: int = 0
    start_day: int = 0
    finish_day: int = 0
    planned_start: Optional[date] = None
    planned_finish: Optional[date] = None
    weight_pct: float = 0.0
    is_critical: bool = False
    activity_count: int = 0
    child_activities: List[str] = field(default_factory=list)

class WBSGenerator:
    """
    Generates a structured 3-level WBS hierarchy with automatic bottom-up roll-up calculations:
    - Level 1: Project Root (Total Project Cost, Duration, Start/Finish, 100% Weight)
    - Level 2: Work Packages (Package Cost, Package Duration, Schedule Windows, Rolled-up Weight %)
    - Level 3: Work Activities (Itemized Activity Scope, Quantities, Rates, and CPM Dates)
    """
    @staticmethod
    def generate_wbs_tree(
        project_title: str,
        activities: List[RoadActivity],
        baseline_activities: Optional[List[RoadActivity]] = None
    ) -> List[WBSNode]:
        base_map = {a.activity_id: a for a in (baseline_activities or activities)}
        
        # 1. Group activities by work package (parent_wbs)
        packages: Dict[str, Dict[str, Any]] = {}
        for act in activities:
            pkg_code = act.parent_wbs
            if pkg_code not in packages:
                packages[pkg_code] = {
                    "name": act.work_package,
                    "description": f"Work Package: {act.work_package}",
                    "activities": []
                }
            packages[pkg_code]["activities"].append(act)

        nodes: List[WBSNode] = []

        # 2. Build Level 1 Project Root Node
        tot_opt_cost = sum(a.cost.total_amount for a in activities)
        tot_base_cost = sum(base_map.get(a.activity_id, a).cost.total_amount for a in activities)
        proj_start_day = min((a.es for a in activities), default=0)
        proj_finish_day = max((a.ef for a in activities), default=0)
        proj_start_dt = min((a.planned_start for a in activities if a.planned_start), default=None)
        proj_finish_dt = max((a.planned_finish for a in activities if a.planned_finish), default=None)
        proj_dur = max(proj_finish_day - proj_start_day, 1)

        root_node = WBSNode(
            wbs_code="1.0",
            parent_wbs="",
            level=1,
            name=project_title,
            description="Complete Project EPC Scope of Works",
            activity_id="ROOT",
            is_work_package=False,
            baseline_cost=tot_base_cost,
            optimized_cost=tot_opt_cost,
            baseline_cost_crore=round(tot_base_cost / 10000000.0, 2),
            optimized_cost_crore=round(tot_opt_cost / 10000000.0, 2),
            duration_days=proj_dur,
            start_day=proj_start_day,
            finish_day=proj_finish_day,
            planned_start=proj_start_dt,
            planned_finish=proj_finish_dt,
            weight_pct=100.0,
            is_critical=True,
            activity_count=len(activities),
            child_activities=[a.activity_id for a in activities]
        )
        nodes.append(root_node)

        # 3. Build Level 2 Work Package Nodes and Level 3 Activity Nodes
        for pkg_code, pkg_info in packages.items():
            pkg_acts = pkg_info["activities"]
            pkg_opt_cost = sum(a.cost.total_amount for a in pkg_acts)
            pkg_base_cost = sum(base_map.get(a.activity_id, a).cost.total_amount for a in pkg_acts)
            pkg_start_day = min((a.es for a in pkg_acts), default=0)
            pkg_finish_day = max((a.ef for a in pkg_acts), default=0)
            pkg_start_dt = min((a.planned_start for a in pkg_acts if a.planned_start), default=None)
            pkg_finish_dt = max((a.planned_finish for a in pkg_acts if a.planned_finish), default=None)
            pkg_dur = max(pkg_finish_day - pkg_start_day, 1)
            pkg_weight = round(sum(a.weight_pct for a in pkg_acts), 2)
            pkg_is_crit = any(a.is_critical for a in pkg_acts)

            pkg_node = WBSNode(
                wbs_code=pkg_code,
                parent_wbs="1.0",
                level=2,
                name=pkg_info["name"],
                description=pkg_info["description"],
                activity_id=f"PKG-{pkg_code}",
                is_work_package=True,
                baseline_cost=pkg_base_cost,
                optimized_cost=pkg_opt_cost,
                baseline_cost_crore=round(pkg_base_cost / 10000000.0, 2),
                optimized_cost_crore=round(pkg_opt_cost / 10000000.0, 2),
                duration_days=pkg_dur,
                start_day=pkg_start_day,
                finish_day=pkg_finish_day,
                planned_start=pkg_start_dt,
                planned_finish=pkg_finish_dt,
                weight_pct=pkg_weight,
                is_critical=pkg_is_crit,
                activity_count=len(pkg_acts),
                child_activities=[a.activity_id for a in pkg_acts]
            )
            nodes.append(pkg_node)

            # Level 3: Individual activities
            for act in pkg_acts:
                b_act = base_map.get(act.activity_id, act)
                act_node = WBSNode(
                    wbs_code=act.wbs_code,
                    parent_wbs=pkg_code,
                    level=3,
                    name=act.name,
                    description=act.description,
                    activity_id=act.activity_id,
                    is_work_package=False,
                    quantity=act.quantity,
                    unit=act.unit,
                    baseline_cost=b_act.cost.total_amount,
                    optimized_cost=act.cost.total_amount,
                    baseline_cost_crore=round(b_act.cost.total_amount / 10000000.0, 2),
                    optimized_cost_crore=round(act.cost.total_amount / 10000000.0, 2),
                    duration_days=act.duration_days,
                    start_day=act.es,
                    finish_day=act.ef,
                    planned_start=act.planned_start,
                    planned_finish=act.planned_finish,
                    weight_pct=act.weight_pct,
                    is_critical=act.is_critical,
                    activity_count=1,
                    child_activities=[act.activity_id]
                )
                nodes.append(act_node)

        return nodes
