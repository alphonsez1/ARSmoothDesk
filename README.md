# ARSmoothDesk

ARSmoothDesk is a Windows application designed specifically for Xreal AR glasses that stabilizes display elements in virtual space while enabling subtle tracking of eye movements. This creates a more natural and comfortable AR experience.

## Overview

ARSmoothDesk creates a "smooth follow" experience for virtual displays, where windows remain anchored in space but subtly adjust their position based on eye movement when your head turns. Unlike traditional implementations, windows only follow when they reach the boundary of your field of view.

This early prototype leverages AirAPI_Windows to access sensor data from Xreal AR glasses.

## Features

- Stabilizes virtual displays in space for a more comfortable viewing experience
- Intelligent boundary-based following: content follows eye movement only when reaching field-of-view boundaries

## Known Issues

As this is an early prototype, there are some known issues:

1. **Mouse Pointer Confusion**: The application runs as a full-screen window on the AR display, which means your mouse pointer can move onto the AR glasses display. This may cause confusion when navigating between screens.

2. **Crash on Close**: Currently, the application crashes when you attempt to close the window. We're working on a fix for this issue.

3. **Connectivity Crash**: The application will crash if AR glasses are not connected at startup.

4. **Initial Repositioning**: Approximately 5 seconds after application start, the window will make a sudden unexpected movement even if your head is stationary.

5. **Screen Mirroring Performance**: The screen mirroring feature currently has suboptimal performance.

## Installation

Currently, you can only run the application from source code:

1. Download and install [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9.0)
2. Install [Visual Studio Code](https://code.visualstudio.com/)
3. Clone or download this repository
4. Open the project folder in VS Code
5. In terminal, run `dotnet restore`.
6. Open `main.cs` in the editor
7. Click the Run button (or press F5) to build and run the application

## Requirements

- Windows 10/11
- Xreal AR glasses (I have only tested Air and Air Ultra)
- .NET 9.0 or higher
- DirectX 11 compatible graphics card

## NuGet Dependencies

This project relies on the following NuGet packages:
- SharpDX (4.2.0)
- SharpDX.Direct3D11 (4.2.0)
- SharpDX.DXGI (4.2.0)

These packages are defined in the `.csproj` file and will be automatically downloaded when running `dotnet restore`.

## Development Status

This project is in an early prototype stage. Expect bugs and limited functionality. Contributions and feedback are welcome!

This project was inspired by [ARMoni](https://www.reddit.com/r/Xreal/comments/1brnean/nebula_alternative_for_win_armoni/), but was developed specifically to create a smooth-follow effect when using AR glasses with a laptop.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE.md) file for details.

Third-party licenses for dependencies can be found in the [Licenses](Licenses/) folder.

## Acknowledgements

- Xreal for their AR glasses
- [AirAPI_Windows](https://github.com/MSmithDev/AirAPI_Windows) for providing the API to interact with Xreal glasses
- [HIDAPI](https://github.com/libusb/hidapi) which is a dependency of AirAPI_Windows
- Contributors and testers who have provided valuable feedback
