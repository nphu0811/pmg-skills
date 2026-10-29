---
name: pmg-diagram-motion
description: Specialized diagram and process motion intelligence for PMG. Animates directional process flows, node-by-node causality, vector connector drawing, hierarchical tree expansions, and data packet throughput in native PowerPoint.
---

# PMG Diagram Motion (`pmg-diagram-motion`)

Specialized motion intelligence for flowcharts, process pipelines, and system architecture diagrams in PowerPoint presentations.

## Core Mental Model

> **Diagram animation must strictly follow causal topology.**
> Never reveal nodes randomly or all at once:
> $$\text{Node A} \longrightarrow \text{Connector AB} \longrightarrow \text{Node B} \longrightarrow \text{Connector BC} \longrightarrow \text{Node C}$$
> - Connectors wipe in the exact direction of causal flow (left-to-right, top-to-bottom).
> - Hierarchical architectures expand from central core roots outward to child modules.
> - Data packet indicators travel along pipelines to demonstrate real-time throughput.

---

## Supported Patterns

| Pattern | Shared Primitive | Native PPTX Effect | Semantic Application |
| :--- | :--- | :--- | :--- |
| **`process_flow`** | `SEQUENCE` | `fly_in` / `appear` | Sequential node appearance following causality. |
| **`connector_drawing`** | `DRAW` | `wipe` (directional) | Connector vector drawn in flow direction. |
| **`hierarchical_expansion`** | `FOCUS` + `STAGGER` | `zoom` + `wipe` | Core root anchors $\to$ child modules cascade. |
| **`data_flow`** | `MOVE` | `motion_path` | Signal packet travels through architecture pipeline. |
| **`cycle_flow`** | `SEQUENCE` + `DRAW` | `appear` + `wipe` | Continuous feedback loop between circular stages. |

---

## Python API Usage

```python
from pmg.motion_planners import DiagramMotionPlanner

planner = DiagramMotionPlanner()
intents = planner.plan_motion({
    "pattern": "process_flow",
    "steps": [
        {"id": "ingest_node", "label": "Data Ingestion"},
        {"id": "process_node", "label": "Neural Processing"},
        {"id": "output_node", "label": "Decision Output"},
    ],
    "connectors": [
        {"id": "conn_1_2", "source": "ingest_node", "direction": "right"},
        {"id": "conn_2_3", "source": "process_node", "direction": "right"},
    ]
})

for intent in intents:
    print(f"[{intent.target_id}] -> {intent.primitive.value} ({intent.semantic_purpose})")
```
