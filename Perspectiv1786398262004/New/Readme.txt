This is another simple 3d object viewer in pure vb, but fit in a nice, clean, and well commented class, then distributed with a tight exe. 

Class explained:

External properties:
	RotateX, RotateY, RotateZ - values to get the angles of orientation of the object.
	TranslateX, TranslateY, TranslateZ - same as above but for position in space of the rendered object

External methods:
	SetRotations, SetTranslations - a method used to set the orientation or position of the object
	RenderObject - a method used to draw the object on the supplied canvas, the internal workings of which is explained in the module
	LoadObject - loads object
		strFileName                   As String     'location of the object file
    		Canvas                        As Object     'object(picturebox) to draw on
                intStyle 		      As Integer    'Style to render object 0=points 
										    1=Solid(Culled) *
										    2=Wireframe(Culled) *
										    3=Wireframe
										    4=Solid
                sngCenterofWorldX 	      As Single     'The origin of the object where it is centered, as this changes the objects 	
                sngCenterofWorldY 	      As Single	    'position is relative to it
                sngCenterofWorldZ 	      As Single
                dblScaleFactor 		      As Double	    'Can be considered zoom but is actually a factor that the coordinates are 							                    'Multiplied by to make the object bigger
                lngSetXRotation 	      As Long	    'Original orientation of the object
                lngSetYRotation 	      As Long
                lngSetZRotation               As Long
                blnZorder 		      As Boolean    'Wether the triangles of the object are sorted by their average z value before they
							    'are rendered. It is necessary with more complex objects(Vanish, Taurus, Teapot).
							    'To see its effect look at these objects in solid culled style, with then without 
							    'using the Z Order procedure.
                blnLight 		      As Boolean    'Wether you want to use lighting effects when rendering the object
                sngLightX 		      As Single	    'Position of the Light source
                sngLightY 		      As Single
                sngLightZ 		      As Single

* Culling, otherwise known as backfacing or backface removal, is a process of rendering where the polygons not facing the camera (ie not visible) are not rendered. This increases rendering speed greatly.


The 3d objects (.odf files) supplied vary in size due to how the were created. The zoom for each is different but can be any value. The recommened zoom for each object is show below to avoid confusion.
	
	Boxish 		0.7
	Card		2.0
	Cube2		2.2
	Cube		2.2
	Sphere2		0.5
	Teapot		0.8
	Tourus		0.8
	Vanish2		1.1
	Vanish3		1.1
	Vanish		1.1

Notes:
	* Compiling the project will increase the speed dramatically
	* Options like Lighting and Z Ordering, while better looking do reduce rendering speed
	
	*
	** While the notes in this Readme and inside the project may provied a simple understanding of 3d computer graphics,
	** THIS PROJECT IN NO WAY CLAIMS TO BE A TUTORIAL. The main tutorial I used to complete this project is listed below
	** I know its program sections are in C but the math is universal. You can also find a 3d math tutorial just about anywhere.
	*

Website
1. http://www.geocities.com/SiliconValley/Horizon/6933/3d.html