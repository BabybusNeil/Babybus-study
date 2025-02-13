' ******************************************************************************  
' C:\Users\Neil-HX-Office\AppData\Local\Temp\swx14224\Macro1.swb - macro recorded on 02/12/25 by BabyBus-PC  
' ******************************************************************************  
' 本程序是用 Visual Basic for Applications (VBA) 语言编写的，主要是为 SolidWorks 自动化设计过程创建宏。它通过记录和编写脚本来完成一些自动化任务，例如选择草图、创建圆、设置尺寸、操作视图等。
' 声明SW应用程序对象  
Dim swApp As Object  
' Dim swApp As Object：声明一个名为 swApp 的变量，其类型为 Object，即可以存储任何类型的对象。
' Dim：用来声明变量（即指定变量名称和类型）。
' Object：在 VBA 中表示一个 对象 类型变量，这个变量可以引用任何类型的对象（如 SolidWorks 应用程序、文件、界面控件等）。
' 对象：在编程中，对象 是具有属性、方法和事件的实体。Dim 是 Dimension 的缩写，起初在早期的 BASIC 编程语言中，Dim 用来指定数组的维度（即数组的大小或形状）。
' 关于Dim的解释：在数组的上下文中，“维度”是指数组的结构（如一维、二维数组等）。随着时间的发展，Dim 被广泛应用于 声明变量，而不仅仅是数组。
' 尽管在现代编程中它不再局限于数组的维度，但 Dim 作为 声明变量 的关键词仍然保留了历史上的命名。它是 Visual Basic 和 VBA 中常用的关键字。
' 声明零件对象  
Dim Part As Object  
Dim boolstatus As Boolean  
Dim longstatus As Long, longwarnings As Long  

' 主宏函数  
Sub main()
' Sub：在 VBA 中，Sub 是 Subroutine 的缩写，表示一个没有返回值的过程或子程序。它是一种功能性的代码块，通常执行某种任务或操作。你可以通过 Call 语句或直接调用子程序的名字来执行这个 Sub。
' Set：在 VBA 中，Set 用来将一个对象赋值给一个对象变量。简单来说，Set 是用来给对象类型的变量赋值的。在 VBA 中，所有的对象类型变量都需要通过 Set 关键字来进行赋值。
' 举例：Set swApp = Application.SldWorks 将 SolidWorks 应用程序对象 赋值给 swApp 变量。这意味着 swApp 现在引用了 SolidWorks 应用程序，可以通过 swApp 访问 SolidWorks 的所有功能和属性。
' Application是 VBA 中的一个预定义对象，它指代 应用程序本身。在 SolidWorks 中，Application.SldWorks 指代当前运行的 SolidWorks 应用程序实例，它是与 SolidWorks 软件交互的入口。
' 划重点：通过 Application.SldWorks，你可以访问 SolidWorks 的各种功能和操作。

    ' 获取当前活动的SolidWorks应用程序对象  赋值给swApp变量
    Set swApp = Application.SldWorks  

    ' 获取当前活动文档（零件）  
    Set Part = swApp.ActiveDoc  
    ' ActiveDoc 表示当前活动的文档，可以是零件、装配体或工程图等。通过 swApp.ActiveDoc，你可以获取对当前打开文档的引用。
    ' Part：Part 是程序员自己给变量起的名字。它并不是一个关键字或保留字，只是一个 自定义的变量名。在这段代码中，Part 是一个变量，用于 引用 当前活动的文档。
    ' 因为 SolidWorks 文档是分为不同类型的（零件、装配体和工程图），Part 在这里假设是一个代表 零件文件（Part）的对象。通常，Part 是指零件文件，但如果你当前打开的是一个装配体或工程图，Part 也会引用那个文档，具体取决于活动文档的类型。

    ' 选择零件中的“上视基准面”（PLANE类型）  
    boolstatus = Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)  

    ' 插入草图  
    Part.SketchManager.InsertSketch True  
    Part.ClearSelection2 True  

    ' 创建一个圆的草图  
    Dim skSegment As Object  
    Set skSegment = Part.SketchManager.CreateCircle(0#, 0#, 0#, 0.034195, -0.018128, 0#)  
    ' 在这行代码中，圆心的位置确实不是 (0, 0, 0)，而是 (0.034195, -0.018128, 0) 应该是操作者建模点不准确导致
    Part.ClearSelection2 True  

    ' 选择圆弧段  
    boolstatus = Part.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0.038317639896463, 0, -5.75977181987972E-03, False, 0, Nothing, 0)  


    ' 为该圆弧段添加尺寸  
    Dim myDisplayDim As Object  
    Set myDisplayDim = Part.AddDimension2(5.95378518644409E-02, 0, -1.39447107218139E-03)  
    Part.ClearSelection2 True  

    ' 设置圆的尺寸为0.1  
    Dim myDimension As Object  
    Set myDimension = Part.Parameter("D1@草图2")  
    myDimension.SystemValue = 0.1  

    ' 显示命名视图  
    Part.ShowNamedView2 "*上下二等角轴测", 8  
    Part.ViewZoomtofit2  

    ' 创建拉伸特征，设置参数  
    Dim myFeature As Object  
    Set myFeature = Part.FeatureManager.FeatureExtrusion2(True, False, False, 6, 0, 0.3, 0.01, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, True, True, True, 0, 0, False)  

    ' 禁用轮廓选择  
    Part.SelectionManager.EnableContourSelection = False  

    ' 缩放并平移视图
    Dim swModelView As Object  
    Set swModelView = Part.ActiveView  
    swModelView.Scale2 = 0.895578094141638  

    ' 平移视图位置  
    Dim swTranslation() As Double  
    ReDim swTranslation(0 To 2) As Double  
    swTranslation(0) = 1.03705859831456E-02  
    swTranslation(1) = 4.00459291993064E-03  
    swTranslation(2) = -5.23664406507014E-03  

    ' 将平移数组转换为MathVector  
    Dim swTranslationVar As Variant  
    swTranslationVar = swTranslation  
    Dim swMathUtils As Object  
    Set swMathUtils = swApp.GetMathUtility()  
    Dim swTranslationVector As MathVector  
    Set swTranslationVector = swMathUtils.CreateVector((swTranslationVar))  

    ' 应用平移到视图  
    swModelView.Translation3 = swTranslationVector  

    ' 接下来的代码反复进行缩放和平移视图的操作，不同的缩放比例和平移量  
    ' 以下部分省略，重复执行类似的缩放和平移操作  

    ' 最后清除选择  
    Part.ClearSelection2 True  

End Sub
