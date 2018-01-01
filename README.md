# ApplyConfigTransformInPlace

Provides a quick, no frills, way to apply any Web or App transform to its respective parent in place.  This tool is meant to speed local development.

Note: it will work for any *.config that follows the pattern [Master].config and the transform is [Master].[anything].config.  Provided that the master and transform are in the same folder path.

How it works:
1. Right click on a transform (ex: Web.Release.config).
2. Click 'Apply Config Transform In Place'
3.  The Visual Studio Extension finds the relative Web.config in the same directory and applies the transform.

![alt text](./ReadMeResources/ApplyTransformInPlace1.png)
