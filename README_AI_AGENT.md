# Rhino Inertia Properties Plugin

## Guides

- [RhinoCommon API Documentation](https://developer.rhino3d.com/api/rhinocommon/)
- [Rhino C++ API Documentation](https://mcneel.github.io/rhino-cpp-api-docs/api/cpp/)
- [Eto.Forms Documentation](https://pages.picoe.ca/docs/api/html/R_Project_EtoForms.htm)
- [Grasshopper API Documentation](https://mcneel.github.io/grasshopper-api-docs/api/grasshopper/html/723c01da-9986-4db2-8f53-6f3a7494df75.htm)

This plugin for Rhino calculates section properties such as moments of inertia, section moduli, radii of gyration, centroid coordinates, and cross-sectional areas based on selected curves and material information.

## Features

- **Material Selection**:  
  Choose a material from a list with densities or add a new material.  
- **Curve Selection**:  
  - "Assign Outline": Select the outline curve.  
  - "Assign Hollows": Select optional hollow curves.  
- **Units**:  
  Display in mm (mm², mm⁴, mm³) or cm (cm², cm⁴, cm³), selectable via a dropdown menu.  
- **Calculations**:  
  - Moments of inertia (Iₓ, Iᵧ)  
  - Section moduli (Wₓ, Wᵧ)  
  - Radii of gyration (iₓ, iᵧ)  
  - Centroid coordinates (x̄, ȳ)  
  - Cross-sectional area (A)  
- **User Interface**:  
  Dockable panel with real-time display of results.

## Installation

1. Download the plugin from [Download Link].  
2. Open Rhino and go to `Tools > PlugInManager`.  
3. Click on `Install` and select the `.rhp` file.  
4. Restart Rhino to activate the plugin.

## Usage

1. Open the panel using the command `ShowSectionPropertiesPanel`.  
2. Select a material from the dropdown list or click "Add Material" to add a new material with name and density.  
3. Click "Assign Outline" and select a closed outline curve in the Rhino document.  
4. Click "Assign Hollows" and select optional hollow curves (if any).  
5. Choose the display unit (mm or cm) from the dropdown menu.  
6. Click "Calculate" to compute the section properties.  
7. The results will be displayed in the panel.

## Mathematical Foundations

The calculations are based on engineering mechanics and RhinoCommon:  
- **Moments of Inertia (Iₓ, Iᵧ)**: ∫y²dA and ∫x²dA over the area.  
- **Section Moduli (Wₓ, Wᵧ)**: Iₓ / y_max and Iᵧ / x_max.  
- **Radii of Gyration (iₓ, iᵧ)**: √(Iₓ / A) and √(Iᵧ / A).  
- **Centroid (x̄, ȳ)**: ∫x dA / A and ∫y dA / A.  
- **Cross-Sectional Area (A)**: Net area (outline minus hollows).  

The curves are analyzed using `AreaMassProperties.Compute`, which automatically accounts for hollows.

## Known Limitations

- Supports only closed, planar curves.  
- Hollow curves must lie within the outline curve.  
- Material density is currently displayed for reference only and not used in calculations.

## Future Enhancements

- Visualization of the centroid or selected curves.  
- Export of results (e.g., as CSV).  
- Integration of additional material properties (e.g., modulus of elasticity).

## Contact

For questions or feedback: [Your Contact Information](mailto:adi.muff@gmail.com).