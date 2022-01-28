<div id="top"></div>
<!--
*** Thanks for checking out the Best-README-Template. If you have a suggestion
*** that would make this better, please fork the repo and create a pull request
*** or simply open an issue with the tag "enhancement".
*** Don't forget to give the project a star!
*** Thanks again! Now go create something AMAZING! :D
-->



<!-- PROJECT SHIELDS -->
<!--
*** I'm using markdown "reference style" links for readability.
*** Reference links are enclosed in brackets [ ] instead of parentheses ( ).
*** See the bottom of this document for the declaration of the reference variables
*** for contributors-url, forks-url, etc. This is an optional, concise syntax you may use.
*** https://www.markdownguide.org/basic-syntax/#reference-style-links
-->



#VBA Raymarcher



<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#prerequisites">Prerequisites</a></li>
        <li><a href="#installation">Installation</a></li>
      </ul>
    </li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#roadmap">Roadmap</a></li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#license">License</a></li>
    <li><a href="#contact">Contact</a></li>
    <li><a href="#acknowledgments">Acknowledgments</a></li>
  </ol>
</details>



<!-- ABOUT THE PROJECT -->
## About The Project

This is a small simplistic Raymarcher implementation based on Microsoft Excel's VBA interface.
The algorithm traces three-dimensional objects using their implicit function definitions. Currently, only spheres and cuboids are supported, but the framework provides an interface for the definition of additional shapes.

The project also includes a small demo that showcases all available features.




<!-- GETTING STARTED -->
## Getting Started

This is a short explanation on how to run the Raymarcher yourself.

### Demo

The folder [demo](demo) contains the Excel document [main.xlsb](demo/main.xlsb). This project can simply be opened in Microsoft Excel for a short demonstration of the features. Please keep in mind that Excel will first require the permission to run VBA code downloaded from unknown sources. 

### Installation

Should you instead want to run the raytracer in your own Excel documents, first create the new Excel document (be careful to select the type "Excel Document WITH Macros"). Then, you will have to import all modules and class modules provided in the [src](src) folder into the new project. You can do this in Excel's VisualBasic screen under `File -> Import File`. The imported modules can now be accessed from any other code file in the project.

## Usage

The `Raytracer` class represents the main actor responsible for the process. Upon creating a new instance of the class, the new object will be populated with default values as defined in [src/class-modules/Raytracer.cls](src/class-modules/Raytracer.cls). At runtime, these default values can be changed to modify the behaviour of the raytracer.
A new raytracer can also be created using the `CreateRaycaster(...)` function provided in [src/class-modules/utils.bas](src/class-modules/Utils.bas).

The main properties that can be modified are 

- `cam`: The camera, from which the tracing process is started. It is defined as an instance of the `ViewerCamera` class, which is a simple data structure defining the viewer's viewing direction, their "up" direction and their position in the world.

- `world`: The observable world traced by the raytracer. It is a collection of shapes (as instances of the `WorldSpaceShape` interface) and lights (as instances of the `Light` class). The world object is responsible for the calculation of the distance between a point and all objects it contains, as is required by the raymarching procedure.

- `far`: The maximum distance from the camera that will be inspected by the raytracer.

- `planeDistance`: The distance between the camera and the view-plane, onto which the world will be projected.

- `pixelWidth / pixelHeight`: The height in pixels for the output image of the procedure.

- `planeWidth / planeHeight`: The dimensions of the view-plane, onto which the world will be projected.

- `backgroundColour`: The colour of the background, given as a VBA Long value, used in case no intersection with the world can be found for some pixel.

After setting up the `Raytracer` object, the method `Raytracer.run()` can be used to start the procedure. The method will return a two-dimensional array containing the colour values of every pixel of the output image. This result can then be displayed in Excel similarly to the approach used in the demo.

### WorldSpaceShape
Instances of the `WorldSpaceShape` interface represent objects in the world. Currently, simple implementations for cuboids and spheres are provided, but additional shapes can be defined using the same interface. Classes implementing this interface must define the properties:

- `specularReflection As Double`
- `diffuseReflection As Double`
- `ambientReflection As Double`
- `shininess As Double`
- `colour As Long`

These properties are used in the illumination computation to determine the colour of the object. The three `reflection` values are factors that determine the influence of specular, diffuse and ambient lighting on this object. The `shininess` property defines the strength of shiny reflection where light othogonally touches the surface. The `colour` value is a colour code for the colour of the object (currently, only even-coloured objects are supported. Multiple different colours cannot be used on the same shape).

Furthermore, the `WorldSpaceShape` interface defines the method `Function Distance(p As Vector3) As Double`. For any given point, this function returns its distance to the shape.

Instances of classes that implement this interface can be passed to the world object using the method `WorldSpace.AddShape(shape as WorldSpaceShape)`.

<!-- LICENSE -->
## License

This project is distributed under the MIT License (See [LICENSE.txt](LICENSE.txt) for more information).
