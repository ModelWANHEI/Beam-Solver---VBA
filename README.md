# Beam-Solver---VBA
A Beam Solver UI in Microsoft Excel that Plots Shear, Moment (Torque), Bending Stress, and Deflection (Both Slope and Magnitude) via the Forward Euler Method. Made with VBA, and Unit Tests are ran through S-Frame testing.

This project was built to practice and study:

- VBA syntax and program structure
- User Defined Types (Type)
- Functions and Subroutines
- Dynamic arrays
- UserForms
- Chart generation through the Excel Object Model
- Error handling and input validation
- Numerical methods (Euler integration)
- Applying engineering mechanics in software

Rather than relying heavily on worksheet formulas, the solver performs the calculations entirely in VBA before presenting the results through the Excel interface.

# Current Features
- Simply-supported beams
- Any number of point loads or applied moments
- Automatic reaction force calculation
- Shear force diagram
- Bending moment diagram
- Bending stress calculation
- Numerical beam deflection using Euler integration
- Interactive chart viewer
- UserForm interface for beam configuration

# Current Assumptions

- Simply-supported beams only
- Rectangular beam cross-sections only
- Point loads only
- Linear elastic material behavior
- Euler-Bernoulli beam theory
- Constant Young's Modulus and cross-section along the beam

# How to Use
Open the workbook.
Enable VBA macros:
1) Right click on the downloaded xlsm and go to properties.
2) Check "Unblock." Apply, OK.
3) Opening the xlsm, click "Enable Macros" in the yellow banner.

Press Configure Beam.
Enter:
1) Beam Length
2) Young's Modulus
3) Beam Width
4) Beam Height
5) Analysis Resolution

Enter each point load using:
1) Position
2) Magnitude

Click Add New Point Load for each load, or Add Moment is you wish to submit an applied moment.
To remove a point, click the row you wish to remove and click "Remove Point Load"/"Remove Moment."
Click Solve Beam.
Use the < and > buttons to cycle between:
- Shear
- Moment
- Stress
- Slope
- Deflection

Reaction forces are displayed directly on the configuration window.

Numerical Method

Internal forces are evaluated at discrete locations along the beam according to the user-selected resolution.

Beam slope and deflection are obtained using forward Euler integration of EIv′′(x) = M(x)

A linear correction is then applied so the computed deflection satisfies the simply-supported boundary condition v(0)=v(L)=0.

While Euler integration is not the highest-order numerical method available, it provides accurate results for sufficiently fine discretizations and was intentionally chosen for its simplicity as a learning exercise.

# Future Improvements

Possible future additions include:

- Uniformly distributed loads (UDLs)
- Additional cross-section geometries
- CSV export
- Better graph customization
- Higher-order numerical integration methods

