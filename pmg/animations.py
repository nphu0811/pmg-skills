"""
pmg.animations
==============
Native PowerPoint Intra-Slide Object Animation & Declarative Timeline Engine (v1.2.0).

Supports native PowerPoint Animation Pane effects via OpenXML (<p:timing>) and COM:
- Entrances: appear, fade, fly_in, wipe, zoom
- Emphasis: scale (grow_shrink), rotate (spin), transparency (opacity), pulse
- Motion Paths: linear, curved Bezier (M ... C ... E), arc/loop
- Exits: disappear, fade_out, fly_out, wipe_out
- Timing & Triggers: delay_ms, duration_ms, with_previous, after_previous, on_click,
  sequential chains, parallel groups, repeat_count, auto_reverse, multiple effects per shape,
  and interactive shape-triggered sequences.
"""

from dataclasses import dataclass, field, asdict
from typing import List, Dict, Any, Optional, Tuple
from lxml import etree

from pmg.models import ObjectMotionSpec

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

NSMAP = {"p": P_NS, "a": A_NS, "r": R_NS}

EFFECT_ALIASES = {
    "fade_in": "fade", "zoom_in": "zoom", "spin": "rotate",
    "grow_shrink": "scale", "linear_motion_path": "motion_path",
    "bezier_motion_path": "motion_path",
}
SUPPORTED_EFFECTS = frozenset((
    "appear", "fade", "fly_in", "wipe", "zoom", "scale", "rotate",
    "transparency", "pulse", "motion_path", "disappear", "fade_out",
    "fly_out", "wipe_out",
))


def normalize_effect(effect_type: str) -> str:
    """Resolve documented aliases; never emit an empty/unknown effect."""
    name = effect_type.strip().lower() if isinstance(effect_type, str) else ""
    name = EFFECT_ALIASES.get(name, name)
    if name not in SUPPORTED_EFFECTS:
        raise ValueError(f"Unsupported native PowerPoint animation effect: {effect_type!r}")
    return name


def iter_native_shapes(shapes):
    """Traverse a native shape hierarchy once, retaining real child shape IDs."""
    for shape in shapes:
        yield shape
        if hasattr(shape, 'shapes'):
            yield from iter_native_shapes(shape.shapes)


@dataclass
class AnimationStep:
    """Declarative specification for a single native PowerPoint animation effect."""
    target_shape_name: str
    target_spid: int
    effect_type: str  # "appear", "fade", "fly_in", "wipe", "zoom", "scale", "rotate", "transparency", "pulse", "motion_path", "disappear", "fade_out", "fly_out", "wipe_out"
    trigger: str = "with_previous"  # "with_previous", "after_previous", "on_click", "on_shape_click"
    duration_ms: int = 800
    delay_ms: int = 0
    direction: str = "up"  # "up", "down", "left", "right"
    motion_path: Optional[str] = None  # e.g., "M 0 0 C 0.1 -0.12 0.22 0.12 0.32 0 E"
    rotation_deg: float = 0.0  # e.g., 360.0 or 180.0
    scale_factor: Tuple[float, float] = (1.0, 1.0)  # e.g., (1.25, 1.25)
    opacity_from: float = 1.0
    opacity_to: float = 1.0
    repeat_count: float = 1.0
    auto_reverse: bool = False
    trigger_shape_spid: Optional[int] = None
    initial_geometry: Tuple[float, float, float, float] = (0.0, 0.0, 1.0, 1.0)  # (x, y, w, h) in inches
    computed_start_ms: int = 0
    computed_end_ms: int = 800
    sequence_index: int = 0

    def __post_init__(self):
        self.effect_type = normalize_effect(self.effect_type)

    @property
    def preset_class(self) -> str:
        if self.effect_type in ("appear", "fade", "fly_in", "wipe", "zoom"):
            return "entr"
        if self.effect_type in ("scale", "rotate", "transparency", "pulse"):
            return "emph"
        if self.effect_type == "motion_path":
            return "path"
        if self.effect_type in ("disappear", "fade_out", "fly_out", "wipe_out"):
            return "exit"
        raise ValueError(f"Unsupported native PowerPoint animation effect: {self.effect_type!r}")

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d["preset_class"] = self.preset_class
        return d


class SlideAnimationTimeline:
    """
    Deterministic intra-slide animation timeline builder and OpenXML <p:timing> compiler.
    """

    def __init__(self, slide_number: int = 1, canvas_width_in: float = 13.333, canvas_height_in: float = 7.5):
        self.slide_number = slide_number
        self.canvas_width_in = canvas_width_in
        self.canvas_height_in = canvas_height_in
        self.steps: List[AnimationStep] = []
        self._shape_geom: Dict[int, Tuple[float, float, float, float]] = {}
        self._shape_type: Dict[int, str] = {}

    def register_shape(self, shape, object_type: str = "shape") -> int:
        if isinstance(shape, (tuple, list)):
            shape = shape[0]
        spid = shape.shape_id
        x = round(shape.left.inches, 3) if hasattr(shape, "left") and shape.left is not None else 0.0
        y = round(shape.top.inches, 3) if hasattr(shape, "top") and shape.top is not None else 0.0
        w = round(shape.width.inches, 3) if hasattr(shape, "width") and shape.width is not None else 1.0
        h = round(shape.height.inches, 3) if hasattr(shape, "height") and shape.height is not None else 1.0
        self._shape_geom[spid] = (x, y, w, h)
        self._shape_type[spid] = object_type
        return spid

    def add_effect(
        self,
        shape,
        effect_type: str,
        trigger: str = "with_previous",
        duration_ms: int = 800,
        delay_ms: int = 0,
        direction: str = "up",
        motion_path: Optional[str] = None,
        rotation_deg: float = 0.0,
        scale_factor: Tuple[float, float] = (1.0, 1.0),
        opacity_from: float = 1.0,
        opacity_to: float = 1.0,
        repeat_count: float = 1.0,
        auto_reverse: bool = False,
        trigger_shape=None,
    ) -> AnimationStep:
        effect_type = normalize_effect(effect_type)
        if trigger not in ("with_previous", "after_previous", "on_click", "on_shape_click"):
            raise ValueError(f"Unsupported animation trigger: {trigger!r}")
        if isinstance(shape, (tuple, list)):
            shape = shape[0]
        if isinstance(trigger_shape, (tuple, list)):
            trigger_shape = trigger_shape[0]
        spid = self.register_shape(shape)
        trig_spid = trigger_shape.shape_id if trigger_shape is not None else None
        geom = self._shape_geom[spid]
        step = AnimationStep(
            target_shape_name=shape.name or f"Shape_{spid}",
            target_spid=spid,
            effect_type=effect_type,
            trigger=trigger,
            duration_ms=max(1, int(duration_ms)),
            delay_ms=max(0, int(delay_ms)),
            direction=direction,
            motion_path=motion_path,
            rotation_deg=rotation_deg,
            scale_factor=scale_factor,
            opacity_from=opacity_from,
            opacity_to=opacity_to,
            repeat_count=repeat_count,
            auto_reverse=auto_reverse,
            trigger_shape_spid=trig_spid,
            initial_geometry=geom,
            sequence_index=len(self.steps) + 1,
        )
        self.steps.append(step)
        self._recompute_schedule()
        return step

    def _recompute_schedule(self) -> None:
        """Deterministically compute slide-relative start and end times (ms) for all steps."""
        current_group_start = 0
        current_group_end = 0
        for idx, step in enumerate(self.steps):
            step.sequence_index = idx + 1
            effective_dur = int(step.duration_ms * max(1.0, step.repeat_count) * (2 if step.auto_reverse else 1))
            if idx == 0:
                step.computed_start_ms = step.delay_ms
                step.computed_end_ms = step.computed_start_ms + effective_dur
                current_group_start = 0
                current_group_end = step.computed_end_ms
            elif step.trigger == "with_previous":
                step.computed_start_ms = current_group_start + step.delay_ms
                step.computed_end_ms = step.computed_start_ms + effective_dur
                current_group_end = max(current_group_end, step.computed_end_ms)
            elif step.trigger in ("after_previous", "on_click", "on_shape_click"):
                # Starts after the previous group finishes (plus optional click wait / delay)
                click_gap = 250 if step.trigger in ("on_click", "on_shape_click") else 0
                current_group_start = current_group_end + click_gap
                step.computed_start_ms = current_group_start + step.delay_ms
                step.computed_end_ms = step.computed_start_ms + effective_dur
                current_group_end = step.computed_end_ms

    def total_animation_duration_ms(self) -> int:
        self._recompute_schedule()
        if not self.steps:
            return 0
        return max(s.computed_end_ms for s in self.steps)

    def to_object_motion_specs(self, slide_abs_start_ms: int = 0, transition_offset_ms: int = 0) -> Tuple[List[ObjectMotionSpec], List[Dict[str, Any]]]:
        """Convert timeline steps into ObjectMotionSpec list and click schedule for ExpectedTimeline."""
        self._recompute_schedule()
        specs: List[ObjectMotionSpec] = []
        clicks: List[Dict[str, Any]] = []
        # Track cumulative position per spid across multiple sequential effects targeting one shape
        live_pos: Dict[int, Tuple[float, float]] = {}
        live_size: Dict[int, Tuple[float, float]] = {}
        for spid, (x, y, w, h) in self._shape_geom.items():
            live_pos[spid] = (x, y)
            live_size[spid] = (w, h)

        for step in self.steps:
            spid = step.target_spid
            init_pos = live_pos.get(spid, (step.initial_geometry[0], step.initial_geometry[1]))
            init_size = live_size.get(spid, (step.initial_geometry[2], step.initial_geometry[3]))
            final_pos = init_pos
            final_size = init_size
            rot_params = {"start_deg": 0.0, "end_deg": step.rotation_deg}
            op_params = {"start_alpha": step.opacity_from, "end_alpha": step.opacity_to}

            if step.preset_class == "entr":
                op_params = {"start_alpha": 0.0, "end_alpha": 1.0}
                if step.effect_type == "fly_in":
                    dx, dy = self._direction_offset(step.direction, entering=True)
                    init_pos = (round(final_pos[0] + dx, 3), round(final_pos[1] + dy, 3))
            elif step.preset_class == "exit":
                op_params = {"start_alpha": 1.0, "end_alpha": 0.0}
                if step.effect_type == "fly_out":
                    dx, dy = self._direction_offset(step.direction, entering=False)
                    final_pos = (round(init_pos[0] + dx, 3), round(init_pos[1] + dy, 3))
            elif step.effect_type == "scale":
                sx, sy = step.scale_factor
                if not step.auto_reverse:
                    final_size = (round(init_size[0] * sx, 3), round(init_size[1] * sy, 3))
                    live_size[spid] = final_size
                else:
                    final_size = (round(init_size[0] * sx, 3), round(init_size[1] * sy, 3))
            elif step.effect_type == "motion_path" and step.motion_path:
                dx_rel, dy_rel = self._parse_motion_path_endpoint(step.motion_path)
                dx_in = dx_rel * self.canvas_width_in
                dy_in = dy_rel * self.canvas_height_in
                final_pos = (round(init_pos[0] + dx_in, 3), round(init_pos[1] + dy_in, 3))
                if not step.auto_reverse:
                    live_pos[spid] = final_pos

            rel_start = transition_offset_ms + step.computed_start_ms
            rel_end = transition_offset_ms + step.computed_end_ms
            abs_start = slide_abs_start_ms + rel_start
            abs_end = slide_abs_start_ms + rel_end

            if step.trigger in ("on_click", "on_shape_click"):
                clicks.append({
                    "sequence_index": step.sequence_index,
                    "target_object": step.target_shape_name,
                    "trigger_type": step.trigger,
                    "slide_relative_ms": rel_start,
                    "absolute_ms": abs_start,
                })

            specs.append(
                ObjectMotionSpec(
                    object_id=step.target_shape_name,
                    slide_number=self.slide_number,
                    object_type=self._shape_type.get(spid, "shape"),
                    initial_position=init_pos,
                    final_position=final_pos,
                    initial_size=init_size,
                    final_size=final_size,
                    rotation_params=rot_params,
                    opacity_params=op_params,
                    motion_path=step.motion_path,
                    start_time_ms=rel_start,
                    duration_ms=step.duration_ms,
                    delay_ms=step.delay_ms,
                    trigger_type=step.trigger,
                    animation_sequence=step.sequence_index,
                    easing=step.effect_type,
                    time_reference="event_triggered" if step.trigger in ("on_click", "on_shape_click") else "slide_relative",
                    absolute_start_ms=abs_start,
                    absolute_end_ms=abs_end,
                )
            )
        return specs, clicks

    @staticmethod
    def _direction_offset(direction: str, entering: bool = True) -> Tuple[float, float]:
        # Offset in inches for Fly In / Fly Out preview calculation
        d = direction.lower()
        sign = 1.0 if entering else -1.0
        if d == "up":
            return (0.0, 2.5 * sign)  # fly in from bottom (upward)
        if d == "down":
            return (0.0, -2.5 * sign)
        if d == "left":
            return (3.0 * sign, 0.0)
        if d == "right":
            return (-3.0 * sign, 0.0)
        return (0.0, 2.0 * sign)

    @staticmethod
    def _parse_motion_path_endpoint(path_str: str) -> Tuple[float, float]:
        """Extract final relative (dx, dy) from an OpenXML path string 'M 0 0 ... x y E'."""
        tokens = [t for t in path_str.replace(",", " ").split() if t not in ("M", "L", "C", "Z", "E", "c", "l", "m", "e")]
        nums = []
        for t in tokens:
            try:
                nums.append(float(t))
            except ValueError:
                pass
        if len(nums) >= 2:
            return (nums[-2], nums[-1])
        return (0.15, 0.0)

    def compile_to_openxml(self, slide) -> etree._Element:
        """
        Compile the declarative timeline into a valid native PowerPoint <p:timing> XML tree
        and attach it to the slide so all effects appear in PowerPoint's native Animation Pane.
        """
        valid_ids = {int(node.get("id")) for node in slide._element.iter(f"{{{P_NS}}}cNvPr")}
        for step in self.steps:
            step.effect_type = normalize_effect(step.effect_type)
            if step.target_spid not in valid_ids:
                raise ValueError(f"Animation target {step.target_spid} does not exist on this slide")
            if step.trigger_shape_spid is not None and step.trigger_shape_spid not in valid_ids:
                raise ValueError(f"Animation trigger {step.trigger_shape_spid} does not exist on this slide")
        self._recompute_schedule()
        sld = slide._element

        # Remove any existing <p:timing>
        existing = sld.find(f"{{{P_NS}}}timing")
        if existing is not None:
            sld.remove(existing)

        if not self.steps:
            return None

        timing = etree.Element(f"{{{P_NS}}}timing", nsmap=NSMAP)
        tnLst = etree.SubElement(timing, f"{{{P_NS}}}tnLst")
        root_par = etree.SubElement(tnLst, f"{{{P_NS}}}par")

        next_id = [1]

        def alloc_id() -> str:
            val = str(next_id[0])
            next_id[0] += 1
            return val

        root_cTn = etree.SubElement(
            root_par,
            f"{{{P_NS}}}cTn",
            id=alloc_id(),
            dur="indefinite",
            restart="never",
            nodeType="tmRoot",
        )
        root_children = etree.SubElement(root_cTn, f"{{{P_NS}}}childTnLst")

        # Main sequence (<p:seq>)
        main_seq = etree.SubElement(root_children, f"{{{P_NS}}}seq", concurrent="1", nextAc="seek")
        main_seq_id = alloc_id()
        main_seq_cTn = etree.SubElement(
            main_seq,
            f"{{{P_NS}}}cTn",
            id=main_seq_id,
            dur="indefinite",
            nodeType="mainSeq",
        )
        main_seq_children = etree.SubElement(main_seq_cTn, f"{{{P_NS}}}childTnLst")

        # Group steps into click groups (each click group contains 1+ after_previous groups,
        # and each after_previous group contains 1+ with_previous effects)
        main_steps = [s for s in self.steps if s.trigger != "on_shape_click"]
        interactive_steps = [s for s in self.steps if s.trigger == "on_shape_click"]

        click_groups: List[List[List[AnimationStep]]] = []
        for idx, step in enumerate(main_steps):
            if idx == 0 or step.trigger == "on_click":
                click_groups.append([[step]])
            elif step.trigger == "after_previous":
                click_groups[-1].append([step])
            else:  # with_previous
                click_groups[-1][-1].append(step)

        grp_counter = [0]
        bld_entries: List[Tuple[int, int]] = []

        for cg_idx, after_groups in enumerate(click_groups):
            click_par = etree.SubElement(main_seq_children, f"{{{P_NS}}}par")
            click_cTn = etree.SubElement(click_par, f"{{{P_NS}}}cTn", id=alloc_id(), fill="hold")
            stCondLst = etree.SubElement(click_cTn, f"{{{P_NS}}}stCondLst")
            etree.SubElement(stCondLst, f"{{{P_NS}}}cond", delay="indefinite")
            # If the very first effect of the click group is automatic (with_previous / after_previous),
            # add onBegin so PowerPoint plays it automatically on slide start!
            first_step = after_groups[0][0]
            if cg_idx == 0 and first_step.trigger in ("with_previous", "after_previous"):
                cond_begin = etree.SubElement(stCondLst, f"{{{P_NS}}}cond", evt="onBegin", delay="0")
                etree.SubElement(cond_begin, f"{{{P_NS}}}tn", val=main_seq_id)

            click_children = etree.SubElement(click_cTn, f"{{{P_NS}}}childTnLst")
            cumulative_after_delay = 0

            for ag_idx, parallel_steps in enumerate(after_groups):
                after_par = etree.SubElement(click_children, f"{{{P_NS}}}par")
                after_cTn = etree.SubElement(after_par, f"{{{P_NS}}}cTn", id=alloc_id(), fill="hold")
                ag_st = etree.SubElement(after_cTn, f"{{{P_NS}}}stCondLst")
                etree.SubElement(ag_st, f"{{{P_NS}}}cond", delay=str(cumulative_after_delay if ag_idx > 0 else 0))
                after_children = etree.SubElement(after_cTn, f"{{{P_NS}}}childTnLst")

                max_group_span = 0
                for p_idx, step in enumerate(parallel_steps):
                    grp_id = grp_counter[0]
                    grp_counter[0] += 1
                    bld_entries.append((step.target_spid, grp_id))

                    if cg_idx > 0 and ag_idx == 0 and p_idx == 0 and step.trigger == "on_click":
                        node_type = "clickEffect"
                    elif ag_idx > 0 and p_idx == 0:
                        node_type = "afterEffect"
                    else:
                        node_type = "withEffect"

                    self._build_effect_par(after_children, step, node_type, grp_id, alloc_id)
                    eff_dur = int(step.duration_ms * max(1.0, step.repeat_count) * (2 if step.auto_reverse else 1))
                    max_group_span = max(max_group_span, step.delay_ms + eff_dur)

                cumulative_after_delay += max_group_span

        # Prev / Next conditions for mainSeq
        prevCondLst = etree.SubElement(main_seq, f"{{{P_NS}}}prevCondLst")
        prevCond = etree.SubElement(prevCondLst, f"{{{P_NS}}}cond", evt="onPrev", delay="0")
        prevTgt = etree.SubElement(prevCond, f"{{{P_NS}}}tgtEl")
        etree.SubElement(prevTgt, f"{{{P_NS}}}sldTgt")

        nextCondLst = etree.SubElement(main_seq, f"{{{P_NS}}}nextCondLst")
        nextCond = etree.SubElement(nextCondLst, f"{{{P_NS}}}cond", evt="onNext", delay="0")
        nextTgt = etree.SubElement(nextCond, f"{{{P_NS}}}tgtEl")
        etree.SubElement(nextTgt, f"{{{P_NS}}}sldTgt")

        # Interactive shape-triggered sequences (if any)
        for ist in interactive_steps:
            trig_spid = ist.trigger_shape_spid or ist.target_spid
            iseq = etree.SubElement(root_children, f"{{{P_NS}}}seq", concurrent="1", nextAc="seek")
            iseq_id = alloc_id()
            iseq_cTn = etree.SubElement(
                iseq, f"{{{P_NS}}}cTn", id=iseq_id, dur="indefinite", evtFilter="cancelBubble", nodeType="interactiveSeq"
            )
            st_i = etree.SubElement(iseq_cTn, f"{{{P_NS}}}stCondLst")
            c_i = etree.SubElement(st_i, f"{{{P_NS}}}cond", evt="onClick", delay="0")
            t_i = etree.SubElement(c_i, f"{{{P_NS}}}tgtEl")
            etree.SubElement(t_i, f"{{{P_NS}}}spTgt", spid=str(trig_spid))

            iseq_children = etree.SubElement(iseq_cTn, f"{{{P_NS}}}childTnLst")
            ipar = etree.SubElement(iseq_children, f"{{{P_NS}}}par")
            icTn = etree.SubElement(ipar, f"{{{P_NS}}}cTn", id=alloc_id(), fill="hold")
            ist_cond = etree.SubElement(icTn, f"{{{P_NS}}}stCondLst")
            etree.SubElement(ist_cond, f"{{{P_NS}}}cond", delay="0")
            ichild = etree.SubElement(icTn, f"{{{P_NS}}}childTnLst")

            ipar2 = etree.SubElement(ichild, f"{{{P_NS}}}par")
            icTn2 = etree.SubElement(ipar2, f"{{{P_NS}}}cTn", id=alloc_id(), fill="hold")
            ist_cond2 = etree.SubElement(icTn2, f"{{{P_NS}}}stCondLst")
            etree.SubElement(ist_cond2, f"{{{P_NS}}}cond", delay="0")
            ichild2 = etree.SubElement(icTn2, f"{{{P_NS}}}childTnLst")

            grp_id = grp_counter[0]
            grp_counter[0] += 1
            bld_entries.append((ist.target_spid, grp_id))
            self._build_effect_par(ichild2, ist, "clickEffect", grp_id, alloc_id)

            inext = etree.SubElement(iseq, f"{{{P_NS}}}nextCondLst")
            inext_c = etree.SubElement(inext, f"{{{P_NS}}}cond", evt="onClick", delay="0")
            inext_t = etree.SubElement(inext_c, f"{{{P_NS}}}tgtEl")
            etree.SubElement(inext_t, f"{{{P_NS}}}spTgt", spid=str(trig_spid))

        # <p:bldLst>
        if bld_entries:
            bldLst = etree.SubElement(timing, f"{{{P_NS}}}bldLst")
            for spid, grp_id in bld_entries:
                etree.SubElement(bldLst, f"{{{P_NS}}}bldP", spid=str(spid), grpId=str(grp_id), animBg="1")

        sld.append(timing)
        return timing

    def _build_effect_par(
        self,
        parent_el: etree._Element,
        step: AnimationStep,
        node_type: str,
        grp_id: int,
        alloc_id,
    ) -> etree._Element:
        eff_par = etree.SubElement(parent_el, f"{{{P_NS}}}par")
        preset_id, preset_class, preset_subtype = self._resolve_preset_attrs(step)

        cTn_attrs = {
            "id": alloc_id(),
            "presetID": str(preset_id),
            "presetClass": preset_class,
            "presetSubtype": str(preset_subtype),
            "fill": "hold",
            "grpId": str(grp_id),
            "nodeType": node_type,
        }
        if step.repeat_count > 1.0:
            cTn_attrs["repeatCount"] = str(int(step.repeat_count * 1000))
        if step.auto_reverse:
            cTn_attrs["autoRev"] = "1"
        if step.effect_type == "motion_path":
            cTn_attrs["accel"] = "50000"
            cTn_attrs["decel"] = "50000"

        eff_cTn = etree.SubElement(eff_par, f"{{{P_NS}}}cTn", **cTn_attrs)
        st = etree.SubElement(eff_cTn, f"{{{P_NS}}}stCondLst")
        etree.SubElement(st, f"{{{P_NS}}}cond", delay=str(step.delay_ms))
        children = etree.SubElement(eff_cTn, f"{{{P_NS}}}childTnLst")

        spid_str = str(step.target_spid)
        dur_str = str(step.duration_ms)

        if step.effect_type == "appear":
            self._append_visibility_set(children, alloc_id, spid_str, "visible", delay_ms=0)

        elif step.effect_type in ("fade", "zoom"):
            self._append_visibility_set(children, alloc_id, spid_str, "visible", delay_ms=0)
            ae = etree.SubElement(
                children,
                f"{{{P_NS}}}animEffect",
                transition="in",
                filter="fade" if step.effect_type == "fade" else "box(out)",
            )
            cBhvr = etree.SubElement(ae, f"{{{P_NS}}}cBhvr")
            etree.SubElement(cBhvr, f"{{{P_NS}}}cTn", id=alloc_id(), dur=dur_str)
            tgt = etree.SubElement(cBhvr, f"{{{P_NS}}}tgtEl")
            etree.SubElement(tgt, f"{{{P_NS}}}spTgt", spid=spid_str)

        elif step.effect_type == "wipe":
            self._append_visibility_set(children, alloc_id, spid_str, "visible", delay_ms=0)
            filter_dir = {"up": "wipe(down)", "down": "wipe(up)", "left": "wipe(right)", "right": "wipe(left)"}.get(
                step.direction.lower(), "wipe(down)"
            )
            ae = etree.SubElement(children, f"{{{P_NS}}}animEffect", transition="in", filter=filter_dir)
            cBhvr = etree.SubElement(ae, f"{{{P_NS}}}cBhvr")
            etree.SubElement(cBhvr, f"{{{P_NS}}}cTn", id=alloc_id(), dur=dur_str)
            tgt = etree.SubElement(cBhvr, f"{{{P_NS}}}tgtEl")
            etree.SubElement(tgt, f"{{{P_NS}}}spTgt", spid=spid_str)

        elif step.effect_type == "fly_in":
            self._append_visibility_set(children, alloc_id, spid_str, "visible", delay_ms=0)
            x_start, y_start = self._fly_formulas(step.direction)
            self._append_ppt_xy_anim(children, alloc_id, spid_str, dur_str, "ppt_x", x_start, "#ppt_x")
            self._append_ppt_xy_anim(children, alloc_id, spid_str, dur_str, "ppt_y", y_start, "#ppt_y")

        elif step.effect_type == "rotate":
            deg = step.rotation_deg if step.rotation_deg != 0.0 else 360.0
            by_val = str(int(deg * 60000))
            ar = etree.SubElement(children, f"{{{P_NS}}}animRot", by=by_val)
            cBhvr = etree.SubElement(ar, f"{{{P_NS}}}cBhvr")
            etree.SubElement(cBhvr, f"{{{P_NS}}}cTn", id=alloc_id(), dur=dur_str, fill="hold")
            tgt = etree.SubElement(cBhvr, f"{{{P_NS}}}tgtEl")
            etree.SubElement(tgt, f"{{{P_NS}}}spTgt", spid=spid_str)
            attrLst = etree.SubElement(cBhvr, f"{{{P_NS}}}attrNameLst")
            attr = etree.SubElement(attrLst, f"{{{P_NS}}}attrName")
            attr.text = "r"

        elif step.effect_type in ("scale", "pulse"):
            if isinstance(step.scale_factor, (int, float)):
                sf = float(step.scale_factor)
                sx, sy = (sf, sf) if sf != 1.0 else (1.2, 1.2)
            elif isinstance(step.scale_factor, (tuple, list)) and len(step.scale_factor) >= 2:
                sx, sy = step.scale_factor[0], step.scale_factor[1]
            else:
                sx, sy = (1.2, 1.2)
            asc = etree.SubElement(children, f"{{{P_NS}}}animScale")
            cBhvr = etree.SubElement(asc, f"{{{P_NS}}}cBhvr")
            etree.SubElement(cBhvr, f"{{{P_NS}}}cTn", id=alloc_id(), dur=dur_str, fill="hold")
            tgt = etree.SubElement(cBhvr, f"{{{P_NS}}}tgtEl")
            etree.SubElement(tgt, f"{{{P_NS}}}spTgt", spid=spid_str)
            etree.SubElement(asc, f"{{{P_NS}}}by", x=str(int(sx * 100000)), y=str(int(sy * 100000)))

        elif step.effect_type == "transparency":
            op_val = f"{step.opacity_to:.2f}"
            p_set = etree.SubElement(children, f"{{{P_NS}}}set")
            cBhvr = etree.SubElement(p_set, f"{{{P_NS}}}cBhvr")
            etree.SubElement(cBhvr, f"{{{P_NS}}}cTn", id=alloc_id(), dur=dur_str, fill="hold")
            tgt = etree.SubElement(cBhvr, f"{{{P_NS}}}tgtEl")
            etree.SubElement(tgt, f"{{{P_NS}}}spTgt", spid=spid_str)
            attrLst = etree.SubElement(cBhvr, f"{{{P_NS}}}attrNameLst")
            attr = etree.SubElement(attrLst, f"{{{P_NS}}}attrName")
            attr.text = "style.opacity"
            to_el = etree.SubElement(p_set, f"{{{P_NS}}}to")
            etree.SubElement(to_el, f"{{{P_NS}}}strVal", val=op_val)

        elif step.effect_type == "motion_path":
            path_str = step.motion_path or "M 0 0 C 0.08 -0.12 0.18 0.12 0.28 0 E"
            am = etree.SubElement(
                children,
                f"{{{P_NS}}}animMotion",
                origin="layout",
                path=path_str,
                pathEditMode="relative",
                ptsTypes="",
            )
            cBhvr = etree.SubElement(am, f"{{{P_NS}}}cBhvr")
            etree.SubElement(cBhvr, f"{{{P_NS}}}cTn", id=alloc_id(), dur=dur_str, fill="hold")
            tgt = etree.SubElement(cBhvr, f"{{{P_NS}}}tgtEl")
            etree.SubElement(tgt, f"{{{P_NS}}}spTgt", spid=spid_str)
            attrLst = etree.SubElement(cBhvr, f"{{{P_NS}}}attrNameLst")
            a1 = etree.SubElement(attrLst, f"{{{P_NS}}}attrName")
            a1.text = "ppt_x"
            a2 = etree.SubElement(attrLst, f"{{{P_NS}}}attrName")
            a2.text = "ppt_y"

        elif step.effect_type == "disappear":
            self._append_visibility_set(children, alloc_id, spid_str, "hidden", delay_ms=0)

        elif step.effect_type in ("fade_out", "wipe_out"):
            filter_val = "fade" if step.effect_type == "fade_out" else "wipe(down)"
            ae = etree.SubElement(children, f"{{{P_NS}}}animEffect", transition="out", filter=filter_val)
            cBhvr = etree.SubElement(ae, f"{{{P_NS}}}cBhvr")
            etree.SubElement(cBhvr, f"{{{P_NS}}}cTn", id=alloc_id(), dur=dur_str)
            tgt = etree.SubElement(cBhvr, f"{{{P_NS}}}tgtEl")
            etree.SubElement(tgt, f"{{{P_NS}}}spTgt", spid=spid_str)
            self._append_visibility_set(children, alloc_id, spid_str, "hidden", delay_ms=max(0, step.duration_ms - 1))

        elif step.effect_type == "fly_out":
            x_end, y_end = self._fly_formulas(step.direction)
            self._append_ppt_xy_anim(children, alloc_id, spid_str, dur_str, "ppt_x", "#ppt_x", x_end)
            self._append_ppt_xy_anim(children, alloc_id, spid_str, dur_str, "ppt_y", "#ppt_y", y_end)
            self._append_visibility_set(children, alloc_id, spid_str, "hidden", delay_ms=max(0, step.duration_ms - 1))

        return eff_par

    @staticmethod
    def _resolve_preset_attrs(step: AnimationStep) -> Tuple[int, str, int]:
        dir_map = {"up": 4, "down": 1, "left": 2, "right": 8}
        sub = dir_map.get((step.direction or "up").lower(), 4)
        et = step.effect_type
        if et == "appear":
            return (1, "entr", 0)
        if et == "fade":
            return (10, "entr", 0)
        if et == "fly_in":
            return (2, "entr", sub)
        if et == "wipe":
            return (22, "entr", sub)
        if et == "zoom":
            return (23, "entr", 16)
        if et == "rotate":
            return (8, "emph", 0)
        if et in ("scale", "pulse"):
            return (6, "emph", 0)
        if et == "transparency":
            return (9, "emph", 0)
        if et == "motion_path":
            return (63, "path", 0)
        if et == "disappear":
            return (1, "exit", 0)
        if et == "fade_out":
            return (10, "exit", 0)
        if et == "fly_out":
            return (2, "exit", sub)
        if et == "wipe_out":
            return (22, "exit", sub)
        raise ValueError(f"Unsupported native PowerPoint animation effect: {et!r}")

    @staticmethod
    def _fly_formulas(direction: Optional[str] = "up") -> Tuple[str, str]:
        d = (direction or "up").lower()
        if d == "up":  # from bottom moving up
            return ("#ppt_x", "1+#ppt_h/2")
        if d == "down":  # from top moving down
            return ("#ppt_x", "0-#ppt_h/2")
        if d == "left":  # from right moving left
            return ("1+#ppt_w/2", "#ppt_y")
        if d == "right":  # from left moving right
            return ("0-#ppt_w/2", "#ppt_y")
        return ("#ppt_x", "1+#ppt_h/2")

    @staticmethod
    def _append_visibility_set(parent_el: etree._Element, alloc_id, spid_str: str, val: str, delay_ms: int = 0) -> None:
        p_set = etree.SubElement(parent_el, f"{{{P_NS}}}set")
        cBhvr = etree.SubElement(p_set, f"{{{P_NS}}}cBhvr")
        cTn = etree.SubElement(cBhvr, f"{{{P_NS}}}cTn", id=alloc_id(), dur="1", fill="hold")
        st = etree.SubElement(cTn, f"{{{P_NS}}}stCondLst")
        etree.SubElement(st, f"{{{P_NS}}}cond", delay=str(delay_ms))
        tgt = etree.SubElement(cBhvr, f"{{{P_NS}}}tgtEl")
        etree.SubElement(tgt, f"{{{P_NS}}}spTgt", spid=spid_str)
        attrLst = etree.SubElement(cBhvr, f"{{{P_NS}}}attrNameLst")
        attr = etree.SubElement(attrLst, f"{{{P_NS}}}attrName")
        attr.text = "style.visibility"
        to_el = etree.SubElement(p_set, f"{{{P_NS}}}to")
        etree.SubElement(to_el, f"{{{P_NS}}}strVal", val=val)

    @staticmethod
    def _append_ppt_xy_anim(
        parent_el: etree._Element,
        alloc_id,
        spid_str: str,
        dur_str: str,
        attr_name: str,
        start_formula: str,
        end_formula: str,
    ) -> None:
        anim = etree.SubElement(parent_el, f"{{{P_NS}}}anim", calcmode="lin", valueType="num")
        cBhvr = etree.SubElement(anim, f"{{{P_NS}}}cBhvr", additive="base")
        etree.SubElement(cBhvr, f"{{{P_NS}}}cTn", id=alloc_id(), dur=dur_str, fill="hold")
        tgt = etree.SubElement(cBhvr, f"{{{P_NS}}}tgtEl")
        etree.SubElement(tgt, f"{{{P_NS}}}spTgt", spid=spid_str)
        attrLst = etree.SubElement(cBhvr, f"{{{P_NS}}}attrNameLst")
        attr = etree.SubElement(attrLst, f"{{{P_NS}}}attrName")
        attr.text = attr_name
        tavLst = etree.SubElement(anim, f"{{{P_NS}}}tavLst")
        tav0 = etree.SubElement(tavLst, f"{{{P_NS}}}tav", tm="0")
        val0 = etree.SubElement(tav0, f"{{{P_NS}}}val")
        etree.SubElement(val0, f"{{{P_NS}}}strVal", val=start_formula)
        tav1 = etree.SubElement(tavLst, f"{{{P_NS}}}tav", tm="100000")
        val1 = etree.SubElement(tav1, f"{{{P_NS}}}val")
        etree.SubElement(val1, f"{{{P_NS}}}strVal", val=end_formula)


def _xpath(el, expr: str):
    try:
        return el.xpath(expr)
    except (TypeError, etree.XPathEvalError):
        return el.xpath(expr, namespaces=NSMAP)


def parse_slide_timing_openxml(slide, slide_number: int = 1, slide_abs_start_ms: int = 0, transition_offset_ms: int = 0) -> Tuple[List[ObjectMotionSpec], List[Dict[str, Any]]]:
    """
    Parse an existing OpenXML <p:timing> tree on a slide and extract its intra-slide
    ObjectMotionSpecs and click schedule.
    """
    sld = slide._element
    timing = sld.find(f"{{{P_NS}}}timing")
    if timing is None:
        return [], []

    shape_by_spid: Dict[int, Any] = {}
    for shp in iter_native_shapes(slide.shapes):
        shape_by_spid[shp.shape_id] = shp

    specs: List[ObjectMotionSpec] = []
    clicks: List[Dict[str, Any]] = []

    effect_nodes = _xpath(
        timing,
        './/p:cTn[@presetClass="entr" or @presetClass="emph" or @presetClass="path" or @presetClass="exit"]',
    )

    current_time_ms = transition_offset_ms
    last_group_start_ms = transition_offset_ms
    last_group_end_ms = transition_offset_ms

    for idx, eff_cTn in enumerate(effect_nodes):
        preset_class = eff_cTn.get("presetClass", "entr")
        preset_id = eff_cTn.get("presetID", "10")
        node_type = eff_cTn.get("nodeType", "withEffect")
        auto_rev = eff_cTn.get("autoRev") == "1"
        rep_raw = eff_cTn.get("repeatCount")
        rep_count = (int(rep_raw) / 1000.0) if rep_raw and rep_raw.isdigit() else 1.0

        # Delay
        delay_ms = 0
        conds = _xpath(eff_cTn, "./p:stCondLst/p:cond[@delay]")
        if conds:
            d_val = conds[0].get("delay", "0")
            if d_val.isdigit():
                delay_ms = int(d_val)

        # Target spid and duration
        sp_tgts = _xpath(eff_cTn, ".//p:spTgt[@spid]")
        if not sp_tgts:
            continue
        spid = int(sp_tgts[0].get("spid"))
        shp = shape_by_spid.get(spid)
        obj_name = shp.name if shp is not None else f"Shape_{spid}"

        dur_ms = 500
        child_ctns = _xpath(eff_cTn, ".//p:cBhvr/p:cTn[@dur]")
        for c in child_ctns:
            d = c.get("dur", "500")
            if d.isdigit() and int(d) > 1:
                dur_ms = int(d)
                break

        eff_span = int(dur_ms * max(1.0, rep_count) * (2 if auto_rev else 1))

        if idx == 0:
            start_ms = current_time_ms + delay_ms
            end_ms = start_ms + eff_span
            last_group_start_ms = current_time_ms
            last_group_end_ms = end_ms
        elif node_type == "withEffect":
            start_ms = last_group_start_ms + delay_ms
            end_ms = start_ms + eff_span
            last_group_end_ms = max(last_group_end_ms, end_ms)
        else:  # afterEffect or clickEffect
            gap = 250 if node_type == "clickEffect" else 0
            last_group_start_ms = last_group_end_ms + gap
            start_ms = last_group_start_ms + delay_ms
            end_ms = start_ms + eff_span
            last_group_end_ms = end_ms

        x = round(shp.left.inches, 3) if shp is not None and hasattr(shp, "left") and shp.left is not None else 0.0
        y = round(shp.top.inches, 3) if shp is not None and hasattr(shp, "top") and shp.top is not None else 0.0
        w = round(shp.width.inches, 3) if shp is not None and hasattr(shp, "width") and shp.width is not None else 1.0
        h = round(shp.height.inches, 3) if shp is not None and hasattr(shp, "height") and shp.height is not None else 1.0

        init_pos = (x, y)
        final_pos = (x, y)
        init_size = (w, h)
        final_size = (w, h)
        rot_params = {"start_deg": 0.0, "end_deg": 0.0}
        op_params = {"start_alpha": 1.0, "end_alpha": 1.0}
        motion_path = None

        if preset_class == "entr":
            op_params = {"start_alpha": 0.0, "end_alpha": 1.0}
        elif preset_class == "exit":
            op_params = {"start_alpha": 1.0, "end_alpha": 0.0}

        anim_rot = _xpath(eff_cTn, ".//p:animRot[@by]")
        if anim_rot:
            by_v = anim_rot[0].get("by", "0")
            if by_v.lstrip("-").isdigit():
                rot_params["end_deg"] = round(int(by_v) / 60000.0, 2)

        anim_scale = _xpath(eff_cTn, ".//p:animScale/p:by")
        if anim_scale:
            bx = int(anim_scale[0].get("x", "100000")) / 100000.0
            by = int(anim_scale[0].get("y", "100000")) / 100000.0
            final_size = (round(w * bx, 3), round(h * by, 3))

        anim_motion = _xpath(eff_cTn, ".//p:animMotion[@path]")
        if anim_motion:
            motion_path = anim_motion[0].get("path")
            dx_rel, dy_rel = SlideAnimationTimeline._parse_motion_path_endpoint(motion_path)
            final_pos = (round(x + dx_rel * 13.333, 3), round(y + dy_rel * 7.5, 3))

        trig_map = {"withEffect": "with_previous", "afterEffect": "after_previous", "clickEffect": "on_click"}
        trig_type = trig_map.get(node_type, "with_previous")

        if trig_type == "on_click":
            clicks.append({
                "sequence_index": idx + 1,
                "target_object": obj_name,
                "trigger_type": "on_click",
                "slide_relative_ms": start_ms,
                "absolute_ms": slide_abs_start_ms + start_ms,
            })

        specs.append(
            ObjectMotionSpec(
                object_id=obj_name,
                slide_number=slide_number,
                object_type="shape",
                initial_position=init_pos,
                final_position=final_pos,
                initial_size=init_size,
                final_size=final_size,
                rotation_params=rot_params,
                opacity_params=op_params,
                motion_path=motion_path,
                start_time_ms=start_ms,
                duration_ms=dur_ms,
                delay_ms=delay_ms,
                trigger_type=trig_type,
                animation_sequence=idx + 1,
                easing=f"{preset_class}_{preset_id}",
                time_reference="event_triggered" if trig_type == "on_click" else "slide_relative",
                absolute_start_ms=slide_abs_start_ms + start_ms,
                absolute_end_ms=slide_abs_start_ms + end_ms,
            )
        )

    return specs, clicks


def verify_animation_pane_com(pptx_path: str) -> Dict[str, Any]:
    """
    Open the PPTX in Microsoft PowerPoint Desktop via COM and inspect the actual
    native Animation Pane (slide.TimeLine.MainSequence and InteractiveSequences)
    to confirm all effects are recognized natively by PowerPoint.
    """
    import os
    abs_path = os.path.abspath(pptx_path)
    result = {
        "verified_via_com": False,
        "slides": {},
        "total_main_sequence_effects": 0,
        "total_interactive_effects": 0,
        "error": None,
    }
    import faulthandler
    fh_enabled = faulthandler.is_enabled()
    if fh_enabled:
        faulthandler.disable()
    app = None
    pres = None
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("PowerPoint.Application")
        pres = app.Presentations.Open(abs_path, True, False, False)
        total_main = 0
        total_inter = 0
        for s_idx in range(1, pres.Slides.Count + 1):
            sl = pres.Slides(s_idx)
            m_cnt = sl.TimeLine.MainSequence.Count
            i_cnt = sl.TimeLine.InteractiveSequences.Count
            eff_details = []
            for e_i in range(1, m_cnt + 1):
                eff = sl.TimeLine.MainSequence(e_i)
                eff_details.append({
                    "index": e_i,
                    "shape_name": eff.Shape.Name,
                    "effect_type": int(eff.EffectType),
                    "trigger_type": int(eff.Timing.TriggerType),
                    "duration_sec": round(float(eff.Timing.Duration), 3),
                    "exit": bool(eff.Exit),
                })
            result["slides"][s_idx] = {
                "main_sequence_count": m_cnt,
                "interactive_sequence_count": i_cnt,
                "effects": eff_details,
            }
            total_main += m_cnt
            total_inter += i_cnt
        result["total_main_sequence_effects"] = total_main
        result["total_interactive_effects"] = total_inter
        result["verified_via_com"] = True
    except Exception as e:
        result["error"] = str(e)
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if fh_enabled:
            faulthandler.enable()
    return result
