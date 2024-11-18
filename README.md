Option Strict Off
Imports System
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.UI
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Drawing

Module ColorMaterial
    Dim theSession As Session = Session.GetSession()
    Dim workPart As Part = theSession.Parts.Work
    Dim theUfSession As UFSession = UFSession.GetUFSession
    Dim theUi As UI = UI.GetUI()
    Dim lw As ListingWindow = theSession.ListingWindow
    Dim listingWindowVisible As Boolean = False ' Trạng thái hiện tại của Listing Window
    Dim listingWindowOpen As Boolean = False ' Cờ theo dõi trạng thái cửa sổ Listing
' Khai báo biến private cho distances
Private distances(2) As Double
Dim KL As String = String.Empty
Dim enteredSize As String = String.Empty


    Sub Main()
        Application.EnableVisualStyles()
        Application.Run(New ColorForm())
    End Sub

    Public Class ColorForm
        Inherits Form

        Private checkBoxes1 As New List(Of CheckBox)()
        Private checkBoxes2 As New List(Of CheckBox)()
        Private colorPanels1 As New List(Of Panel)()
        Private colorPanels2 As New List(Of Panel)()
        Private totalCheckBox As CheckBox
    Private TextBoxModel As TextBox
    Private ComboBoxPartNo As ComboBox
    Private ComboBoxPartName As ComboBox
    Private ComboBoxQuantity As ComboBox
    Private ComboBoxRemark As ComboBox
    Private TextBoxSize As TextBox
    Private TextBoxMass As TextBox

        Private colorMapSolid As New Dictionary(Of String, Integer) From {
            {"Al6061", 123},
		  {"S45C", 125},
            {"SKD11", 142},
		  {"SUS304", 162},
            {"POM", 201},
            {"Bakelite", 78},
            {"PC", 103},
            {"Acrylic", 31},
            {"Peek", 109},
            {"Urethan 90", 6},
            {"Urethan 70", 159},
            {"Brass", 42},
            {"Sheet Metal", 80},
            {"Standard Part", 26}
        }

        Private colorMapFace As New Dictionary(Of String, Integer) From {
            {"Modification", 186},
            {"Precision Face", 181},
            {"Datum Face", 103},
            {"3D Face", 108}
        }

        Public Sub New()
            ' KHỞI TẠO FORM
            Me.Text = "Select Material Color"
            Me.Size = New Drawing.Size(350, 480)
            Me.StartPosition = FormStartPosition.Manual ' Đặt thủ công để có thể tùy chỉnh vị trí
            Me.TopMost = True ' Ensure the form is always on top
            Me.BackColor = Color.White ' Set background color to white
            ' TÍNH TOÁN VỊ TRÍ HIỂN THỊ FORM
            Dim rightmostScreen As Screen = Screen.AllScreens(0)
            For Each scr As Screen In Screen.AllScreens
                If scr.Bounds.Right > rightmostScreen.Bounds.Right Then
                    rightmostScreen = scr
                End If
            Next
            Dim formX As Integer = rightmostScreen.Bounds.Right - Me.Width
            Dim formY As Integer = (rightmostScreen.Bounds.Height - Me.Height) \ 2 '
            formX = formX - (Me.Width \ 1.2)
            Me.Location = New System.Drawing.Point(formX, formY)

            Dim columnOffset As Integer = 170
            Dim yPos As Integer = 20
		Dim thirdColumnOffset As Integer = 170

		'TẠO CỘT CHECK BOX THỨ NHẤT (BODY)
            For Each kvp As KeyValuePair(Of String, Integer) In colorMapSolid
                Dim checkBox As New CheckBox()
                checkBox.Text = kvp.Key
                checkBox.Location = New Drawing.Point(20, yPos)
                checkBox.ForeColor = Color.Black ' Set text color to black
                checkBox.AutoSize = True ' Automatically size the checkbox
                ' Create a color panel next to the checkbox
                Dim colorPanel As New Panel()
                colorPanel.BackColor = GetColorFromCode(kvp.Value) ' Set panel background color
                colorPanel.Location = New Drawing.Point(120, yPos - 2) ' Align it vertically with the checkbox
                colorPanel.Size = New Drawing.Size(20, 20) ' Set size for the color panel
                AddHandler checkBox.CheckedChanged, AddressOf SolidCheckBox_CheckedChanged
                Me.Controls.Add(checkBox)
                Me.Controls.Add(colorPanel)
                checkBoxes1.Add(checkBox)
                colorPanels1.Add(colorPanel)
                yPos += 30
            Next

            ' TẠO CỘT CHECK BOX THỨ 2 (FACE)
            yPos = 20
            For Each kvp As KeyValuePair(Of String, Integer) In colorMapFace
                Dim checkBox As New CheckBox()
                checkBox.Text = kvp.Key
                checkBox.Location = New Drawing.Point(columnOffset, yPos)
                checkBox.ForeColor = Color.Black ' Set text color to black
                checkBox.AutoSize = True ' Automatically size the checkbox

                ' Create a color panel next to the checkbox
                Dim colorPanel As New Panel()
                colorPanel.BackColor = GetColorFromCode(kvp.Value) ' Set panel background color
                colorPanel.Location = New Drawing.Point(columnOffset + 120, yPos - 2) ' Align it vertically with the checkbox
                colorPanel.Size = New Drawing.Size(20, 20) ' Set size for the color panel

                ' Add event handler for checkbox change
                AddHandler checkBox.CheckedChanged, AddressOf FaceCheckBox_CheckedChanged

                Me.Controls.Add(checkBox)
                Me.Controls.Add(colorPanel)

                checkBoxes2.Add(checkBox)
                colorPanels2.Add(colorPanel)

                yPos += 30
            Next
' TẠO CỘT ATTRIBUTE

    ' Các tùy chọn ComboBox
    Dim partNameOptions As New List(Of String) From {
        "Part", "Pin", "Shaft", "Base", "Plate", "Block", "Guide", "Bracket", "Housing",
        "Holder", "Clamp", "Handle", "Spacer", "Stopper", "Link", "Tool", "Master",
        "Gauge", "Cover", "Frame", "Assy"
    }

    Dim remarkOptions As New List(Of String) From {
        "Black-Anod", "White-Anod", "Hard-Anod", "Cr-Anod", "Ni-Anod", "Cr/Ni Anod",
        "DLC Coating", "Barrel", "HrC40±2", "HrC54±2", "HrC60±2"
    }

    Dim partNoOptions As New List(Of String) From {
        "101", "102", "103", "104", "105", "106", "107", "108", "109", "110",
        "111", "112", "113", "114", "115", "116", "117", "118", "119", "120"
    }

    Dim quantityOptions As New List(Of String) From {
        "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
    }

    ' Vị trí ban đầu của label

    ' Tạo label và Model cho "Model"
    Dim labelModel As New Label()
    labelModel.Text = "Model"
    labelModel.Location = New Drawing.Point(thirdColumnOffset, yPos)
    labelModel.ForeColor = Color.Black
    labelModel.AutoSize = True
    Me.Controls.Add(labelModel)

    TextBoxModel = New TextBox()
    TextBoxModel.Location = New Drawing.Point(thirdColumnOffset + 60, yPos - 3)
    TextBoxModel.Size = New Drawing.Size(80, 20)
    Me.Controls.Add(TextBoxModel)

    ' Cập nhật vị trí yPos cho ComboBox và label tiếp theo
    yPos += 30

    ' Tạo label và ComboBox cho "Part No"
    Dim labelPartNo As New Label()
    labelPartNo.Text = "Part No"
    labelPartNo.Location = New Drawing.Point(thirdColumnOffset , yPos)
    labelPartNo.ForeColor = Color.Black
    labelPartNo.AutoSize = True
    Me.Controls.Add(labelPartNo)

    ComboBoxPartNo = New ComboBox()
    ComboBoxPartNo.Location = New Drawing.Point(thirdColumnOffset + 60, yPos - 3)
    ComboBoxPartNo.Size = New Drawing.Size(80, 20)
    ComboBoxPartNo.DropDownStyle = ComboBoxStyle.DropDown
    ComboBoxPartNo.Items.AddRange(partNoOptions.ToArray()) ' Gán các tùy chọn cho ComboBox Part No
    Me.Controls.Add(ComboBoxPartNo)

' Add event handler for ComboBoxPartNo selection change
AddHandler ComboBoxPartNo.SelectedIndexChanged, AddressOf ComboBoxPartNo_SelectedIndexChanged


    ' Cập nhật vị trí yPos cho ComboBox và label tiếp theo
    yPos += 30

' Tạo label và ComboBox cho "Part Name"
Dim labelPartName As New Label()
labelPartName.Text = "Part Name"
labelPartName.Location = New Drawing.Point(thirdColumnOffset, yPos)
labelPartName.ForeColor = Color.Black
labelPartName.AutoSize = True
Me.Controls.Add(labelPartName)

ComboBoxPartName = New ComboBox()
ComboBoxPartName.Location = New Drawing.Point(thirdColumnOffset + 60, yPos - 3)
ComboBoxPartName.Size = New Drawing.Size(80, 20)
ComboBoxPartName.DropDownStyle = ComboBoxStyle.DropDown
ComboBoxPartName.Items.AddRange(partNameOptions.ToArray()) ' Gán các tùy chọn cho ComboBox Part Name
ComboBoxPartName.AutoCompleteMode = AutoCompleteMode.SuggestAppend
ComboBoxPartName.AutoCompleteSource = AutoCompleteSource.ListItems
Me.Controls.Add(ComboBoxPartName)

    ' Cập nhật vị trí yPos cho ComboBox và label tiếp theo
    yPos += 30

    ' Tạo label và ComboBox cho "Quantity"
    Dim labelQuantity As New Label()
    labelQuantity.Text = "Quantity"
    labelQuantity.Location = New Drawing.Point(thirdColumnOffset, yPos)
    labelQuantity.ForeColor = Color.Black
    labelQuantity.AutoSize = True
    Me.Controls.Add(labelQuantity)

    ComboBoxQuantity = New ComboBox()
    ComboBoxQuantity.Location = New Drawing.Point(thirdColumnOffset + 60, yPos - 3)
    ComboBoxQuantity.Size = New Drawing.Size(80, 20)
    ComboBoxQuantity.DropDownStyle = ComboBoxStyle.DropDown
    ComboBoxQuantity.Items.AddRange(quantityOptions.ToArray()) ' Gán các tùy chọn cho ComboBox Quantity
    Me.Controls.Add(ComboBoxQuantity)

    ' Cập nhật vị trí yPos cho ComboBox và label tiếp theo
    yPos += 30

' Tạo label và ComboBox cho "Remark"
Dim labelRemark As New Label()
labelRemark.Text = "Remark"
labelRemark.Location = New Drawing.Point(thirdColumnOffset, yPos)
labelRemark.ForeColor = Color.Black
labelRemark.AutoSize = True
Me.Controls.Add(labelRemark)

ComboBoxRemark = New ComboBox()
ComboBoxRemark.Location = New Drawing.Point(thirdColumnOffset + 60, yPos - 3)
ComboBoxRemark.Size = New Drawing.Size(80, 20)
ComboBoxRemark.DropDownStyle = ComboBoxStyle.DropDown
ComboBoxRemark.Items.AddRange(remarkOptions.ToArray()) ' Gán các tùy chọn cho ComboBox Remark
ComboBoxRemark.AutoCompleteMode = AutoCompleteMode.SuggestAppend
ComboBoxRemark.AutoCompleteSource = AutoCompleteSource.ListItems
Me.Controls.Add(ComboBoxRemark)


    ' Cập nhật vị trí yPos cho ComboBox và label tiếp theo
    yPos += 30

    ' Tạo label và TextBox cho "Size"
    Dim labelSize As New Label()
    labelSize.Text = "Size"
    labelSize.Location = New Drawing.Point(thirdColumnOffset, yPos)
    labelSize.ForeColor = Color.Black
    labelSize.AutoSize = True
    Me.Controls.Add(labelSize)

    TextBoxSize = New TextBox()
    TextBoxSize.Location = New Drawing.Point(thirdColumnOffset + 60, yPos - 3)
    TextBoxSize.Size = New Drawing.Size(80, 20)
    Me.Controls.Add(TextBoxSize)

    ' Cập nhật vị trí yPos cho TextBox
    yPos += 30

    ' Tạo label và TextBox cho "Mass"
    Dim labelMass As New Label()
    labelMass.Text = "Mass"
    labelMass.Location = New Drawing.Point(thirdColumnOffset, yPos)
    labelMass.ForeColor = Color.Black
    labelMass.AutoSize = True
    Me.Controls.Add(labelMass)

    TextBoxMass = New TextBox()
    TextBoxMass.Location = New Drawing.Point(thirdColumnOffset + 60, yPos - 3)
    TextBoxMass.Size = New Drawing.Size(80, 20)
    Me.Controls.Add(TextBoxMass)

    ' Cập nhật vị trí yPos cho TextBox
    yPos += 30

    ' Add the "Group" checkbox below the text boxes
    Dim groupCheckBox As New CheckBox()
    groupCheckBox.Text = "Group"
    groupCheckBox.Location = New Drawing.Point(thirdColumnOffset, yPos)
    groupCheckBox.ForeColor = Color.Black
    groupCheckBox.AutoSize = True
    groupCheckBox.Checked = True ' Default to checked
    Me.Controls.Add(groupCheckBox)

    ' Add the "Body Name" checkbox below the text boxes
    Dim BodyNameCheckBox As New CheckBox()
    BodyNameCheckBox.Text = "Body Name"
    BodyNameCheckBox.Location = New Drawing.Point(thirdColumnOffset + 65, yPos)
    BodyNameCheckBox.ForeColor = Color.Black
    BodyNameCheckBox.AutoSize = True
    BodyNameCheckBox.Checked = True ' Default to checked
    Me.Controls.Add(BodyNameCheckBox)

            ' Add the Toggle Listing Window button
            Dim toggleButton As New Button
            toggleButton.Text = "MEASURE"
            toggleButton.Size = New Drawing.Size(145, 30)
            toggleButton.Location = New Drawing.Point(170, 375)
            AddHandler toggleButton.Click, AddressOf ToggleListingWindow
            Me.Controls.Add(toggleButton)

' Thêm checkbox Total dưới button MEASURE
totalCheckBox = New CheckBox()
totalCheckBox.Text = "Total Mass"
totalCheckBox.Size = New Drawing.Size(100, 30)
totalCheckBox.Location = New Drawing.Point(170, 405) ' Đặt vị trí dưới button MEASURE
AddHandler totalCheckBox.CheckedChanged, AddressOf TotalCheckBox_CheckedChanged
Me.Controls.Add(totalCheckBox)


            ' Hook up the FormClosing event to handle closing restrictions
            AddHandler Me.FormClosing, AddressOf ColorForm_FormClosing
        End Sub

'_________________________________________XỬ LÝ FORM, CHECKBOX________________________________________________________________________
'HÀM NGĂN ĐÓNG FORM (TRÁNH LỖI)
        Private Sub ColorForm_FormClosing(sender As Object, e As FormClosingEventArgs)
            If isSelectingObjects Then
                ' Nếu đang trong quá trình chọn đối tượng, không cho phép đóng cửa sổ
                MessageBox.Show("Hmmm ~ Đóng cửa sổ Selection trước nhé!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                e.Cancel = True ' Ngừng quá trình đóng cửa sổ
            End If
        End Sub

'HÀM VÔ HIỆU HÓA- KÍCH HOẠT CHECK BOX
        Private Sub ToggleCheckBoxes(ByVal enable As Boolean)
            ' Disable or enable all checkboxes
            For Each checkBox As CheckBox In checkBoxes1
                checkBox.Enabled = enable
            Next
            For Each checkBox As CheckBox In checkBoxes2
                checkBox.Enabled = enable
            Next
        End Sub

' BIẾN THEO DÕI TRẠNG THÁI SELECT
        Private isSelectingObjects As Boolean = False
 	   Private selectedBodies As New List(Of Body)

' Sự kiện khi checkbox "Total" được thay đổi trạng thái
Private Sub TotalCheckBox_CheckedChanged(sender As Object, e As EventArgs)
    If totalCheckBox.Checked Then
        ' Nếu checkbox "Total" được chọn, tính cộng dồn khối lượng
        totalMass = 0 ' Khởi tạo lại biến totalMass
    End If
End Sub

'''DEL
Private Sub ComboBoxPartNo_SelectedIndexChanged(sender As Object, e As EventArgs)
    ' Clear the other ComboBoxes
    ComboBoxPartName.SelectedIndex = -1
    ComboBoxQuantity.SelectedIndex = -1
    ComboBoxRemark.SelectedIndex = -1

    ' Clear the TextBoxes
    TextBoxSize.Clear()
    TextBoxMass.Clear()

    ' Optionally uncheck checkboxes if needed
    totalCheckBox.Checked = False
    ' You can add other checkboxes here if needed.
End Sub


'___________________________________________FUNCTION FACE_________________________________________

'SỰ KIỆN CHECKBOX FACES : Gán Color
Private Sub FaceCheckBox_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
    Dim checkBox As CheckBox = CType(sender, CheckBox)
    If checkBox.Checked Then
        Dim colorCode As Integer = colorMapFace(checkBox.Text)
        ' Vô hiệu hóa tất cả các checkbox khi bắt đầu chọn đối tượng
        ToggleCheckBoxes(False)

        ' Gán màu và attribute cho các mặt đã chọn
        SelectFacesAndApplyColor(colorCode)
        
        checkBox.Checked = False '
    End If
End Sub

'PHƯƠNG THỨC GÁN COLOR CHO FACES ĐƯỢC LỰA CHỌN
        Private Sub SelectFacesAndApplyColor(colorCode As Integer)
            Dim selectedFaces() As NXObject = Nothing
            Dim response As Selection.Response

            ' Chọn mặt sử dụng phương thức SelectFaces từ code tham khảo
            If SelectFaces(selectedFaces) Then
                ' Duyệt qua các mặt đã chọn và áp dụng màu
                For Each selectedFace As NXObject In selectedFaces
                    Dim displayableObject As DisplayableObject = CType(selectedFace, DisplayableObject)
                    Dim displayModification As DisplayModification = theSession.DisplayManager.NewDisplayModification()
                    With displayModification
                        .ApplyToAllFaces = True
                        .NewColor = colorCode
                        .Apply(New DisplayableObject() {displayableObject})
                    End With
                    displayModification.Dispose()
                Next
            End If
        End Sub

' ___________________________________________ FUNCTION BODY ___________________________________________

'SỰ KIỆN CHECKBOX BODY:
Private Sub SolidCheckBox_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
    Dim checkBox As CheckBox = CType(sender, CheckBox)
    If checkBox.Checked Then
        ' Vô hiệu hóa tất cả các checkbox khi bắt đầu chọn đối tượng
        ToggleCheckBoxes(False)

        ' Gán màu và attribute cho các body đã chọn và gán mật độ
        AssignColorAttributesAndDensityToSelectedBodies(colorMapSolid(checkBox.Text), checkBox.Text)


        ' Sau khi tính toán xong, gọi hàm gán thuộc tính
        AssignAttributesToWorkPart()

        ' Uncheck the checkbox after color assignment
        checkBox.Checked = False

    End If
End Sub



' ___________________________________________ ASSIGN ATTRIBUTES ___________________________________________


' Gán thuộc tính cho WorkPart
Private Sub AssignAttributesToWorkPart()
    ' Lấy phiên làm việc và Part hiện tại
    Dim theSession As Session = Session.GetSession()
    Dim workPart As Part = theSession.Parts.Work
    Dim WP(0) As NXObject
    WP(0) = workPart

    ' Tạo AttributePropertiesBuilder để xây dựng thuộc tính
    Dim attributeBuilder As AttributePropertiesBuilder = theSession.AttributeManager.CreateAttributePropertiesBuilder(workPart, WP, AttributePropertiesBuilder.OperationType.None)

    ' Lấy giá trị từ các điều khiển trên Form
    Dim selectedPartNo As String = ComboBoxPartNo.Text
    If ComboBoxPartNo.SelectedItem IsNot Nothing Then
        selectedPartNo = ComboBoxPartNo.SelectedItem.ToString()
    End If

    Dim selectedPartName As String = ComboBoxPartName.Text
    If ComboBoxPartName.SelectedItem IsNot Nothing Then
        selectedPartName = ComboBoxPartName.SelectedItem.ToString()
    End If

    Dim selectedQuantity As String = ComboBoxQuantity.Text
    If ComboBoxQuantity.SelectedItem IsNot Nothing Then
        selectedQuantity = ComboBoxQuantity.SelectedItem.ToString()
    End If

    Dim selectedRemark As String = ComboBoxRemark.Text
    If ComboBoxRemark.SelectedItem IsNot Nothing Then
        selectedRemark = ComboBoxRemark.SelectedItem.ToString()
    End If

    If Not String.IsNullOrEmpty(TextBoxSize.Text) Then
        ' Nếu đã có giá trị trong TextBoxSize, sử dụng giá trị đó
        enteredSize = TextBoxSize.Text
Else
    enteredSize = " "
End If

If Not String.IsNullOrEmpty(TextBoxMass.Text) Then
    ' Nếu đã có giá trị trong TextBoxMass, sử dụng giá trị đó
    KL = TextBoxMass.Text
Else
    ' Nếu không có giá trị, gán KL là một chuỗi rỗng
    KL = " "
End If

    ' Kết hợp danh sách CheckBox
    Dim allCheckBoxes As New List(Of CheckBox)(checkBoxes1)
    allCheckBoxes.AddRange(checkBoxes2)

    Dim selectedMaterial As String = Nothing
    For Each checkBox As CheckBox In allCheckBoxes
        If checkBox.Checked Then
            selectedMaterial = checkBox.Text
            Exit For
        End If
    Next

    ' Gán giá trị cho Category (sử dụng Part No làm Category nếu có)
    attributeBuilder.Category = If(selectedPartNo, "Unknown")


    ' Gán vào thuộc tính "Size"
    If Not String.IsNullOrEmpty(selectedPartNo) Then
        attributeBuilder.Title = String.Format("{0} Size", selectedPartNo)
        attributeBuilder.StringValue = enteredSize
        attributeBuilder.CreateAttribute()
    End If

    If Not String.IsNullOrEmpty(KL) Then
        attributeBuilder.Title = String.Format("{0} Mass", selectedPartNo)
        attributeBuilder.StringValue = KL  ' Gán giá trị KL vào thuộc tính "Mass"
        attributeBuilder.CreateAttribute()
    End If

  ' Kiểm tra nếu Part No có giá trị mới tạo các Attribute
    If Not String.IsNullOrEmpty(selectedPartNo) Then
        ' Thiết lập các thuộc tính
        If Not String.IsNullOrEmpty(selectedPartNo) Then
            attributeBuilder.Title = String.Format("{0} Part No", selectedPartNo)
            attributeBuilder.StringValue = selectedPartNo
            attributeBuilder.CreateAttribute()
        End If

        If Not String.IsNullOrEmpty(selectedPartName) Then
            attributeBuilder.Title = String.Format("{0} Part Name", selectedPartNo)
            attributeBuilder.StringValue = selectedPartName
            attributeBuilder.CreateAttribute()
        End If

        If Not String.IsNullOrEmpty(selectedMaterial) Then
            attributeBuilder.Title = String.Format("{0} Material", selectedPartNo)
            attributeBuilder.StringValue = selectedMaterial
            attributeBuilder.CreateAttribute()
        End If

        If Not String.IsNullOrEmpty(selectedQuantity) Then
            attributeBuilder.Title = String.Format("{0} Quantity", selectedPartNo)
            attributeBuilder.StringValue = selectedQuantity
            attributeBuilder.CreateAttribute()
        End If

        If Not String.IsNullOrEmpty(selectedRemark) Then
            attributeBuilder.Title = String.Format("{0} Remark", selectedPartNo)
            attributeBuilder.StringValue = selectedRemark
            attributeBuilder.CreateAttribute()
        End If   
 End If

    ' Hoàn tất và hủy Attribute Builder
    attributeBuilder.Commit()
    attributeBuilder.Destroy()
End Sub

'PHƯƠNG THỨC GÁN CÁC THUỘC TÍNH CHO BODY: Color, Attribute, Density
Private Sub AssignColorAttributesAndDensityToSelectedBodies(ByVal colorCode As Integer, ByVal materialName As String)
    Dim selectedObjectsArray() As NXObject = Nothing
    If SelectBodies(selectedObjectsArray) Then
        ' Duyệt qua các đối tượng đã chọn
        Dim displayableObjects(selectedObjectsArray.Length - 1) As DisplayableObject
        For i As Integer = 0 To selectedObjectsArray.Length - 1
            displayableObjects(i) = CType(selectedObjectsArray(i), DisplayableObject)
        Next

        ' Gán màu cho các body đã chọn
        Dim displayModification As DisplayModification = theSession.DisplayManager.NewDisplayModification()
        With displayModification
            .ApplyToAllFaces = True
            .NewColor = colorCode
            .Apply(displayableObjects)
        End With
        displayModification.Dispose()

        ' Gán Attribute cho từng Body
        For Each selectedObject As NXObject In selectedObjectsArray
            If TypeOf selectedObject Is Body Then
                Dim body As Body = CType(selectedObject, Body)
                AssignAttributesToBody(body, materialName)
    	   ' Gán mật độ cho body
                Dim density As Double = GetDensityForMaterial(materialName)
                AssignDensityAndMeasureProperties(body, density)
            End If
        Next

        ' Bật lại các checkbox sau khi gán màu, vật liệu và đo đạc xong
        ToggleCheckBoxes(True)

    End If
End Sub

'______________________________________________________CHỌN ĐỐI TƯỢNG_____________________________________________
'HÀM CHỌN ĐỐI TƯỢNG BODY
        Function SelectBodies(ByRef selectedObjects As NXObject()) As Boolean
            ' Chức năng chọn body và trả về kết quả
            ' Đánh dấu là đang chọn đối tượng
            isSelectingObjects = True
            Dim ui As UI = NXOpen.UI.GetUI()
            Dim message As String = "Chọn Faces. Có thể chọn nhiều Faces 1 lúc trong Part hoặc Assembly"
            Dim title As String = "SELECT BODIES"
            Dim scope As Selection.SelectionScope = Selection.SelectionScope.AnyInAssembly
            Dim keepHighlighted As Boolean = False
            Dim includeFeatures As Boolean = False
            Dim response As Selection.Response
            Dim selectionAction As Selection.SelectionAction = Selection.SelectionAction.ClearAndEnableSpecific

            ' Tạo selection mask cho các body
            Dim selectionMaskArray(1) As Selection.MaskTriple
            selectionMaskArray(0) = CreateMask(UFConstants.UF_component_type, UFConstants.UF_component_subtype, UFConstants.UF_UI_SEL_FEATURE_SOLID_BODY)
            selectionMaskArray(1) = CreateMask(UFConstants.UF_solid_type, 0, UFConstants.UF_UI_SEL_FEATURE_SOLID_BODY)

            response = ui.SelectionManager.SelectObjects(message, title, scope, selectionAction, includeFeatures, keepHighlighted, selectionMaskArray, selectedObjects)

            ' Sau khi chọn xong, không còn trong chế độ chọn đối tượng
            isSelectingObjects = False

            ' Kích hoạt lại các checkbox khi quá trình chọn xong
            ToggleCheckBoxes(True)

            Return response <> Selection.Response.Cancel AndAlso response <> Selection.Response.Back
        End Function

' Tạo Mask Triple để lọc solid
        Function CreateMask(type As Integer, subtype As Integer, solidBodySubtype As Integer) As Selection.MaskTriple
            Dim mask As Selection.MaskTriple
            With mask
                .Type = type
                .Subtype = subtype
                .SolidBodySubtype = solidBodySubtype
            End With
            Return mask
        End Function


'HÀM CHỌN ĐỐI TƯỢNG FACE
        Function SelectFaces(ByRef selectedObjects() As NXObject) As Boolean
            ' Chức năng chọn mặt và trả về kết quả
            ' Đánh dấu là đang chọn đối tượng
            isSelectingObjects = True
            Dim ui As UI = UI.GetUI()
            Dim message As String = "Chọn Faces. Có thể chọn nhiều Faces 1 lúc trong Part hoặc Assembly"
            Dim title As String = "SELECT FACES"
            Dim scope As Selection.SelectionScope = Selection.SelectionScope.AnyInAssembly
            Dim keepHighlighted As Boolean = False
            Dim includeFeatures As Boolean = False
            Dim response As Selection.Response
            Dim selectionAction As Selection.SelectionAction = Selection.SelectionAction.ClearAndEnableSpecific

            ' Tạo mask để chỉ chọn mặt
            Dim selectionMaskArray(1) As Selection.MaskTriple
            selectionMaskArray(0) = CreateMask(UFConstants.UF_component_type, UFConstants.UF_component_subtype, UFConstants.UF_UI_SEL_FEATURE_ANY_FACE)
            selectionMaskArray(1) = CreateMask(UFConstants.UF_face_type, 0, UFConstants.UF_UI_SEL_FEATURE_ANY_FACE)

            response = ui.SelectionManager.SelectObjects(message, title, scope, selectionAction, includeFeatures, keepHighlighted, selectionMaskArray, selectedObjects)

            ' Sau khi chọn xong, không còn trong chế độ chọn đối tượng
            isSelectingObjects = False

            ' Kích hoạt lại các checkbox khi quá trình chọn xong
            ToggleCheckBoxes(True)

            Return response <> Selection.Response.Cancel AndAlso response <> Selection.Response.Back
        End Function

'___________________________________________CÁC THUẬT TOÁN _______________________________________

'HÀM LẤY GIÁ TRỊ DENSITY
Private Function GetDensityForMaterial(ByVal materialName As String) As Double
    ' Bảng ánh xạ vật liệu với mật độ
    Dim densityMap As New Dictionary(Of String, Double) From {
        {"Al6061", 2711},
        {"SUS304", 7928},
        {"POM", 1420},
        {"S45C", 7928},
        {"SKD11", 7928},
        {"Bakelite", 1250},
        {"PC", 1200},
        {"Peek", 1310},
        {"Acrylic", 1150},
        {"Urethan 90", 1200},
        {"Urethan 70", 1200},
        {"Brass", 8409},
        {"Sheet Metal", 7928},
        {"Standard Part", 1000}
    }
    ' Kiểm tra nếu vật liệu có trong bảng ánh xạ
    If densityMap.ContainsKey(materialName) Then
        Return densityMap(materialName)
    Else
    ' Nếu ko có giá trị nào nhận được, trả giá trị về mặc định trước đó
        Return 0
    End If
End Function


'PHƯƠNG THỨC GÁN ATTRIBUTE CHO BODY
Sub AssignAttributesToBody(ByVal body As Body, ByVal materialName As String)
    ' Lấy phiên bản session và work part
    Dim workPart As Part = theSession.Parts.Work

    ' Thiết lập Undo Mark
    Dim markId1 As NXOpen.Session.UndoMarkId = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Assign Attribute to Body")

    ' Tạo và cấu hình AttributePropertiesBuilder
    Dim attributePropertiesBuilder As NXOpen.AttributePropertiesBuilder = theSession.AttributeManager.CreateAttributePropertiesBuilder(workPart, New NXOpen.NXObject() {body}, NXOpen.AttributePropertiesBuilder.OperationType.None)
    attributePropertiesBuilder.IsArray = False
    attributePropertiesBuilder.DataType = NXOpen.AttributePropertiesBaseBuilder.DataTypeOptions.String
    attributePropertiesBuilder.Category = "Body Information"
    attributePropertiesBuilder.Title = "Body Material"
    attributePropertiesBuilder.StringValue = materialName ' Gán giá trị Material là tên của vật liệu

    ' Commit và cập nhật
    attributePropertiesBuilder.Commit()

    ' Xóa Undo Marks và hoàn thành
    theSession.DeleteUndoMark(theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Invisible, "Solid Body Attributes"), Nothing)
    theSession.UpdateManager.DoUpdate(theSession.GetNewestUndoMark(NXOpen.Session.MarkVisibility.Visible))
    theSession.SetUndoMarkName(theSession.GetNewestUndoMark(NXOpen.Session.MarkVisibility.Visible), "Solid Body Attributes")

    ' Dọn dẹp
    attributePropertiesBuilder.Destroy()
End Sub

        ' Toggle Listing Window visibility
Private Sub ToggleListingWindow(sender As Object, e As EventArgs)
    If Not listingWindowVisible Then
        lw.SelectDevice(ListingWindow.DeviceType.Window, "")
        lw.Open()
        lw.WriteLine("Function is started. Select body to measure...")
        lw.Close() 
        listingWindowVisible = True
        listingWindowOpen = True ' Đặt cờ listingWindowOpen thành True khi cửa sổ Listing mở
    Else
        ' Nếu Listing Window đã được mở rồi, đóng và giải phóng nó
        theUfSession.Ui.ExitListingWindow()
        listingWindowVisible = False
        listingWindowOpen = False ' Đặt cờ listingWindowOpen thành False khi cửa sổ Listing đóng
    End If
End Sub

'TÍNH TOÁN THUỘC TÍNH VẬT LÝ

' Biến toàn cục để lưu trữ khối lượng cộng dồn
Dim totalMass As Double = 0

' Cập nhật phương thức tính toán khối lượng với kiểm tra checkbox "Total"
Private Sub AssignDensityAndMeasureProperties(ByVal body As Body, ByVal density As Double)
    ' Kiểm tra xem cửa sổ Listing có mở không
    If Not listingWindowVisible Then
        ' Nếu cửa sổ Listing không mở, không thực hiện tính toán và không hiển thị kết quả
        Exit Sub
    End If

    'Gán khối lượng riêng cho body trước khi thực hiện các phép đo
    Dim solidDensity1 As NXOpen.GeometricAnalysis.SolidDensity = theSession.Parts.Work.AnalysisManager.CreateSolidDensityObject()
    solidDensity1.Units = NXOpen.GeometricAnalysis.SolidDensity.UnitsType.KilogramsPerCubicMeters
    solidDensity1.Density = density ' Gán mật độ cho body

    ' Gán mật độ cho body
    solidDensity1.Solids.Add(body)

    ' Cam kết thay đổi
    solidDensity1.Commit()

    ' Hủy đối tượng SolidDensity sau khi sử dụng
    solidDensity1.Destroy()

    ' Bước 2: Tính toán khối lượng
    Dim bodies As New List(Of Body) From {body}
    Dim myMeasure As MeasureManager = theSession.Parts.Display.MeasureManager()

    ' Tạo đối tượng Đo
    Dim massUnits(4) As Unit
    massUnits(1) = theSession.Parts.Display.UnitCollection.GetBase("Volume")
    massUnits(2) = theSession.Parts.Display.UnitCollection.GetBase("Mass")
    massUnits(3) = theSession.Parts.Display.UnitCollection.GetBase("Length")

    ' Tạo đối tượng MeasureBodies để tính toán khối lượng
    Dim mb As MeasureBodies = myMeasure.NewMassProperties(massUnits, 0.9999, bodies.ToArray())
    mb.InformationUnit = MeasureBodies.AnalysisUnit.CustomUnit

      ' Tính toán Bounding Box
        Dim csys As NXOpen.Tag = NXOpen.Tag.Null
        Dim min_corner(2) As Double
        Dim directions(2, 2) As Double
        Dim distances(2) As Double
        Dim ufs As UFSession = UFSession.GetUFSession()
        ufs.Csys.AskWcs(csys)
        ufs.Modl.AskBoundingBoxExact(bodies(0).Tag, csys, min_corner, directions, distances)

    ' Nếu checkbox "Total" được chọn, cộng dồn khối lượng
    If totalCheckBox.Checked Then
        totalMass += mb.Mass
    End If

    Dim allCheckBoxes As New List(Of CheckBox)(checkBoxes1)
    allCheckBoxes.AddRange(checkBoxes2)

    Dim selectedMaterial As String = Nothing
    For Each checkBox As CheckBox In allCheckBoxes
        If checkBox.Checked Then
            selectedMaterial = checkBox.Text
            Exit For
        End If
    Next

    ' Hiển thị khối lượng trong cửa sổ Listing
    Dim lw As ListingWindow = theSession.ListingWindow()
    lw.Open()
        '
    If totalCheckBox.Checked Then
        ' Nếu "Total" được chọn, hiển thị khối lượng cộng dồn
        lw.WriteLine("Add Body Material:"& SelectedMaterial)
        lw.WriteLine("TOTAL MASS: " & totalMass.ToString("F3") & " kg")
    Else
        ' Nếu không, hiển thị khối lượng của body hiện tại
        lw.WriteLine("Material:"& SelectedMaterial)
        lw.WriteLine("Body Mass: " & mb.Mass.ToString("F3") & " kg")
        lw.WriteLine("Volume: " & mb.Volume.ToString("F0") & " mm^3")
        lw.WriteLine("Centroid: " & mb.Centroid.X.ToString("F3") & " x " & mb.Centroid.Y.ToString("F3") & " x " & mb.Centroid.Z.ToString("F3")&" (XYZ-Dim from WCS)")
        lw.WriteLine("Size: " & distances(0).ToString("F2") & " x " & distances(1).ToString("F2") & " x " & distances(2).ToString("F2")&" (Dim WCS)")
    End If

    lw.WriteLine("**********************")
    lw.Close()

    If String.IsNullOrEmpty(TextBoxSize.Text) Then
 ' Gán giá trị KL vào TextBoxMass
TextBoxSize.Text = enteredSize ' Gán giá trị KL vào TextBoxMass
    End If

If String.IsNullOrEmpty(TextBoxSize.Text) Then
    ' Sắp xếp mảng distances theo thứ tự giảm dần
    Array.Sort(distances)
    Array.Reverse(distances)
    
    ' Định dạng lại TextBoxSize.Text theo thứ tự sắp xếp
    TextBoxSize.Text = String.Format("{0} x {1} x {2}", distances(0).ToString("F0"), distances(1).ToString("F0"), distances(2).ToString("F0"))
End If

    If String.IsNullOrEmpty(TextBoxMass.Text) Then
        TextBoxMass.Text = KL ' Gán giá trị KL vào TextBoxMass
    End If

            If String.IsNullOrEmpty(TextBoxMass.Text) Then
                TextBoxMass.Text = mb.Mass.ToString("F3") & " kg"
            End If

End Sub

        ' Chuyển đổi mã màu thành màu hiển thị
        Private Function GetColorFromCode(colorCode As Integer) As Color
            Select Case colorCode
                Case 123 : Return Color.Gray
                Case 125 : Return Color.Peru
                Case 142 : Return Color.Teal
                Case 162 : Return Color.SaddleBrown
                Case 201 : Return Color.Black
                Case 78  : Return Color.Orange
                Case 103 : Return Color.DodgerBlue
                Case 31  : Return Color.Aqua
                Case 109  : Return Color.HotPink
                Case 6   : Return Color.Yellow
                Case 159 : Return Color.Gray
                Case 42  : Return Color.Gold
                Case 80  : Return Color.Beige
                Case 26  : Return Color.PowderBlue
                Case 186 : Return Color.Red
                Case 181 : Return Color.DeepPink
                Case 108 : Return Color.Green
                Case Else : Return Color.Black
            End Select
        End Function

        ' Return the unload option
        Public Function GetUnloadOption(ByVal dummy As String) As Integer
            Return NXOpen.UF.UFConstants.UF_UNLOAD_IMMEDIATELY
        End Function
    End Class
End Module






















