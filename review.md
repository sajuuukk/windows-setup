# Review of `setup_w11.ps1` for 3D Modeling & Developer Workstations

## Executive Summary
The script provides a solid foundation for a high-performance, low-distraction environment. It correctly prioritizes system responsiveness, networking performance, and developer tooling (Git, SSH, GPG). However, **critical conflicts** exist regarding input handling for 3D software, and some window management tweaks may negatively impact productivity on large displays typical of these workstations.

## Suitability for 3D Modeling & IDEs

### ✅ Pros
*   **Stability**: Enabling `SeparateProcess` for Explorer is excellent. If a shell extension or heavy asset crashes Explorer, it won't take down the entire desktop environment—vital when working with unstable 3D plugins or large file previews.
*   **Network Performance**: Disabling `NetworkThrottlingIndex` benefits the transfer of large assets (textures, models, builds) across local network shares or NAS devices.
*   **Update Safety**: The script disables automatic reboots (`NoAutoRebootWithLoggedOnUsers`). This is crucial for 3D artists running overnight renders, preventing Windows Update from killing a 20-hour render job.
*   **Developer Readiness**: The automated setup of `ssh-agent`, Git, and GPG keys significantly reduces the "Time to Hello World" for a new machine.

### ⚠️ Cons & Risks
*   **CRITICAL: Input Conflict (AltDrag)**
    *   The script installs and auto-starts `AltDrag`, which binds window movement to `Alt + Left Click`.
    *   **Conflict**: Almost all industry-standard 3D software (Maya, Blender, Unity, Unreal Engine, Substance Painter) uses `Alt + Left Click` for viewport rotation/tumbling.
    *   **Impact**: This makes the workstation unusable for 3D work out-of-the-box. The user would have to kill the process or reconfigure the hotkeys immediately.
*   **Window Management**:
    *   The script disables "Snap Assist" and automatic window arrangement.
    *   **Impact**: Developers and artists often use large, high-resolution monitors (or multi-monitor setups). Disabling Windows Snap features (Win+Arrow keys, corner snapping) hinders the ability to quickly arrange reference images, code, and viewports side-by-side.

## Detailed Analysis

### Performance & Registry Tweaks
*   `MenuShowDelay` (0) and `WaitToKillServiceTimeout` (2000): Excellent for perceived responsiveness.
*   `HideFileExt` (0) and `Hidden` (1): Essential for developers to see file types and config files (`.git`, `.vscode`).
*   `DisableThumbnailsOnNetworkFolders`: Good for performance when browsing heavy asset libraries on a NAS.

### Bloatware Removal
*   The removal list is generally safe, but removing `Microsoft.Windows.Photos` leaves the system without a default image viewer.
*   **Recommendation**: Ensure a replacement (like IrfanView, XnView, or PureRef) is installed, as artists frequently need to review texture files and reference images.

### Security (HardeningKitty)
*   Running HardeningKitty in `Audit` mode is safe and provides good visibility into system security without breaking development tools that might rely on specific legacy permissions.

## Recommendations

1.  **Disable or Configure AltDrag**:
    *   *Option A*: Do not install/start AltDrag by default.
    *   *Option B*: Configure AltDrag to use `Win + Left Click` or a different modifier to avoid colliding with standard DCC (Digital Content Creation) navigation schemes.
2.  **Re-enable Window Snapping**:
    *   Consider commenting out or removing the registry tweaks that disable "Snap Assist" and "Joint Resize", as these are productivity boosters for multitasking.
3.  **Image Viewer Replacement**:
    *   If `Microsoft.Windows.Photos` is removed, consider adding a step to install a lightweight viewer via `winget` (e.g., `WinGet install XnView.MP`).
