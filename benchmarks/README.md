# PMG Animation Benchmark Suite (24-case baseline)

The **PMG Animation Benchmark Suite** evaluates PMG's native PowerPoint animation generation, OpenXML `<p:timing>` compilation, Morph choreography, reverse engineering, visibility-aware validation, and 3-loop self-healing across **24 canonical benchmark cases**.

## Benchmark Groups (24 Cases)

### Group A — Single-Object Animations (Cases 01–08)
1. **Case 01 (`case_01_fade_in`):** Native Fade In entrance (`presetID="10"`, `<p:set>` visibility transition).
2. **Case 02 (`case_02_fade_out`):** Native Fade Out exit (`presetClass="exit"`).
3. **Case 03 (`case_03_fly_in_out`):** Directional Fly In and Fly Out choreography.
4. **Case 04 (`case_04_wipe`):** Directional Wipe entrance (`presetID="22"`).
5. **Case 05 (`case_05_rotation`):** 360-degree native Spin emphasis (`<p:animRot>`).
6. **Case 06 (`case_06_scale_grow_shrink`):** 135% native Grow/Shrink emphasis (`<p:animScale>`).
7. **Case 07 (`case_07_linear_motion_path`):** Horizontal linear native motion path (`<p:animMotion>`).
8. **Case 08 (`case_08_curved_bezier_motion_path`):** Cubic Bezier orbital arc motion path (`M 0 0 C ... E`).

### Group B — Multi-Object Choreography (Cases 09–14)
9. **Case 09 (`case_09_simultaneous`):** Three objects animating simultaneously in parallel.
10. **Case 10 (`case_10_sequential`):** Three objects cascading sequentially with stagger delays.
11. **Case 11 (`case_11_with_previous`):** Synchronized `With Previous` title and zoom-in panel.
12. **Case 12 (`case_12_after_previous`):** Chained `After Previous` Fade followed by directional Wipe.
13. **Case 13 (`case_13_on_click`):** Auto intro card followed by interactive `On Click` reveal.
14. **Case 14 (`case_14_multiple_effects_single_object`):** Multiple sequential effects (`Fade In` + `Motion Path` + `Pulse`) on a single object.

### Group C — Advanced PMG Motion (Cases 15–20)
15. **Case 15 (`case_15_morph_position`):** Cross-slide 60 FPS Morph position translation.
16. **Case 16 (`case_16_morph_scale_rotation`):** Cross-slide Morph scale expansion and rotation.
17. **Case 17 (`case_17_traveling_object_across_slides`):** 3-slide continuous traveling entity (`!!Traveling_Satellite`).
18. **Case 18 (`case_18_camera_zoom_simulation`):** Wide establishing shot to 220% macro push-in camera zoom.
19. **Case 19 (`case_19_multi_plane_parallax`):** 3-plane parallax speed ratios (`BG 0.7x`, `Mid 1.0x`, `FG 1.4x`).
20. **Case 20 (`case_20_complex_overlapping_timeline`):** 4-track overlapping Fade, Fly In, Wipe, and Motion Path.

### Group D — Real-World Failure & Visibility Cases (Cases 21–24)
21. **Case 21 (`case_21_hero_image_visibility`):** Hero Earth image (`!!Earth_Grand`) with sibling `Oval 13` decorative outline.
22. **Case 22 (`case_22_title_text_visibility`):** Animated `!!HeroTitle`, `!!HeroSubtitle`, and summary card visibility.
23. **Case 23 (`case_23_multiple_animated_images`):** Three planetary hero images entering with independent timing.
24. **Case 24 (`case_24_failure_detection_and_repair`):** Negative + Auto-Healing test verifying broken disappearing Earth/Title XML is flagged as `FAIL` before being healed to `PASS` with Fade In animations preserved.

## CLI Usage

```bash
pmg benchmark run --all
pmg benchmark run --case 21
pmg benchmark report
```

Use `--worker static` for structural checks. Those cases remain NOT_VERIFIED for
real playback, even when the structural benchmark assertion succeeds. A critical
package/compatibility/playback defect blocks orchestrator certification regardless
of aggregate metrics. Compatible technical fallback is inside the same maximum
three repair attempts; the benchmark runner never creates another repair loop.

The [continued roadmap](../docs/development-roadmap.md) extends this existing suite
with decomposition, shared graphs, cinematic and creative cases in their owning
releases. It does not create another benchmark infrastructure or claim subjective
creative quality from code tests.

