# PMG Animation Benchmark Suite (44 Canonical Cases)

The **PMG Animation Benchmark Suite** evaluates PMG's native PowerPoint animation generation, OpenXML `<p:timing>` compilation, Morph choreography, reverse engineering, visibility-aware validation, and 3-loop self-healing across **44 canonical benchmark cases** (expanded in v1.8.0).

## Benchmark Groups (44 Cases across Groups A..H)

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

### Group E — Advanced Motion Paths & Dynamics (Cases 25–29)
25. **Case 25 (case_25_orbital_bezier_loop):** Closed orbital motion path with dynamic curve approximation.
26. **Case 26 (case_26_staggered_radial_cascade):** 8-node radial blossom expanding outward with sequential stagger.
27. **Case 27 (case_27_multi_segment_waypoint_path):** Linear waypoint trajectory traversing multiple waypoints.
28. **Case 28 (case_28_continuous_micro_pulse_and_spin):** Dual continuous emphasis: pulsing scale + 360-deg rotation.
29. **Case 29 (case_29_reverse_oscillating_path):** Auto-reversing ping-pong linear trajectory with repeat count.

### Group F — Massive Scene Hierarchies & Layered Parallax (Cases 30–34)
30. **Case 30 (case_30_massive_hierarchy_50_nodes):** Single slide with 50 coordinated vector nodes.
31. **Case 31 (case_31_massive_hierarchy_100_nodes):** Single slide with 100 independently addressable objects.
32. **Case 32 (case_32_massive_hierarchy_200_nodes):** Extreme stress: 200 nodes arranged in a constellation grid.
33. **Case 33 (case_33_six_layer_deep_parallax):** Multi-plane parallax with distinct velocity ratios (0.2x to 1.8x).
34. **Case 34 (case_34_counter_tracking_hud_overlay):** Camera moves right while HUD element counter-tracks to lock on screen.

### Group G — Decomposed Asset Choreography & Reassembly (Cases 35–39)
35. **Case 35 (case_35_decomposed_particle_scatter_gather):** Decomposed particles scatter and converge back into position.
36. **Case 36 (case_36_multi_strip_shutter_reveal):** 10 vertical slices staggered revealing background visual.
37. **Case 37 (case_37_radial_sector_blade_rotation):** 6 radial sectors opening like a mechanical iris with rotation.
38. **Case 38 (case_38_vector_polygon_morph_reassembly):** OpenXML vector polygons morphing across slides.
39. **Case 39 (case_39_frosted_glass_depth_stack):** Multi-tier acrylic glass cards with varying opacity and glowing borders.

### Group H — End-to-End Cinematic Storyboard & Motion Director (Cases 40–44)
40. **Case 40 (case_40_long_intra_slide_sequence_15s):** 15-second multi-act intra-slide timeline with 3 causal phases.
41. **Case 41 (case_41_storyboard_to_motion_director_cinematic_hero):** Full StoryboardScene compiled with MotionDirector.
42. **Case 42 (case_42_thematic_traveling_entity_4_scenes):** Traveling hero object traversing 4 distinct scenes with scale and rotation.
43. **Case 43 (case_43_multi_target_synchronized_parallel):** 12 shapes animating in synchronized lockstep.
44. **Case 44 (case_44_clean_powerpoint_open_com_verification):** Composite slide with 20 shapes and parallel animations.

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

