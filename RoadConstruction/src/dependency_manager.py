"""
Dependency Manager Module for validating CPM network consistency and precedence relationships.
"""
from typing import List, Dict, Set
from .models import RoadActivity

class DependencyManager:
    """Validates activity precedence, checks for cycles and missing predecessors."""
    
    @staticmethod
    def validate_network(activities: List[RoadActivity]) -> List[str]:
        errors = []
        act_ids = {a.activity_id for a in activities}

        # Check unique IDs
        if len(act_ids) != len(activities):
            errors.append("Duplicate Activity IDs detected in network.")

        # Check predecessors exist
        adj: Dict[str, List[str]] = {a.activity_id: [] for a in activities}
        for act in activities:
            for pred in act.predecessors:
                if pred not in act_ids:
                    errors.append(f"Activity {act.activity_id} references missing predecessor: {pred}")
                else:
                    adj[pred].append(act.activity_id)

        # Check for cycles using DFS
        visited = {}  # 0: unvisited, 1: visiting, 2: visited
        for act_id in act_ids:
            visited[act_id] = 0

        def dfs(node: str, path: List[str]):
            visited[node] = 1
            for neighbor in adj.get(node, []):
                if visited[neighbor] == 1:
                    errors.append(f"Circular dependency detected: {' -> '.join(path + [neighbor])}")
                elif visited[neighbor] == 0:
                    dfs(neighbor, path + [neighbor])
            visited[node] = 2

        for act_id in act_ids:
            if visited[act_id] == 0:
                dfs(act_id, [act_id])

        return errors
