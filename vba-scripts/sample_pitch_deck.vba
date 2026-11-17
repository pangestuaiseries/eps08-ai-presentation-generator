Sub GenerateConsultingDeck()
    ' ==============================================================================
    ' AI PRESENTATION ARCHITECT — AUTOMATED CONSULTING DECK GENERATOR
    ' Created by Pangestu Rahmat N (@pangesturahmatn)
    ' ==============================================================================
    
    Dim pptApp As PowerPoint.Application
    Dim pptPres As PowerPoint.Presentation
    Dim mySlide As PowerPoint.Slide
    Dim myShape As PowerPoint.Shape
    
    Set pptPres = ActivePresentation
    
    ' SLIDE 1: COVER
    Set mySlide = pptPres.Slides.Add(1, ppLayoutCustom)
    mySlide.Background.Fill.Solid
    mySlide.Background.Fill.ForeColor.RGB = RGB(15, 23, 42) ' Dark Navy
    
    Set myShape = mySlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 180, 800, 200)
    With myShape.TextFrame.TextRange
        .Text = "TRANSFORMASI OTOMASI AI & WEBHOOK 2026" & vbCrLf & _
                "Meningkatkan Kecepatan Respon Operasional hingga 10x Lipat"
        .Font.Name = "Calibri"
        .Font.Size = 36
        .Font.Bold = msoTrue
        .Font.Color.RGB = RGB(255, 255, 255)
    End With
    
    ' SLIDE 2: PROBLEM STATEMENT
    Set mySlide = pptPres.Slides.Add(2, ppLayoutCustom)
    mySlide.Background.Fill.Solid
    mySlide.Background.Fill.ForeColor.RGB = RGB(248, 250, 252) ' Clean White
    
    Set myShape = mySlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 80, 60, 850, 60)
    With myShape.TextFrame.TextRange
        .Text = "PROBLEM: Bottleneck Manual Menguras 40% Jam Kerja Tim"
        .Font.Name = "Calibri"
        .Font.Size = 28
        .Font.Bold = msoTrue
        .Font.Color.RGB = RGB(220, 38, 38)
    End With
    
    Set myShape = mySlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 80, 150, 850, 350)
    With myShape.TextFrame.TextRange
        .Text = "• Rekap Manual Memakan Waktu: Data transaksi harus dipindahkan manual satu per satu." & vbCrLf & vbCrLf & _
                "• Response Time Lambat: Rata-rata pelanggan menunggu 2 jam untuk balasan informasi stok." & vbCrLf & vbCrLf & _
                "• Human Error Tinggi: 12% pesanan salah catat akibat komunikasi chat grup yang tertumpuk."
        .Font.Name = "Calibri"
        .Font.Size = 20
        .Font.Color.RGB = RGB(51, 65, 85)
    End With
    
    ' SLIDE 3: STRATEGIC SOLUTION
    Set mySlide = pptPres.Slides.Add(3, ppLayoutCustom)
    mySlide.Background.Fill.Solid
    mySlide.Background.Fill.ForeColor.RGB = RGB(248, 250, 252)
    
    Set myShape = mySlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 80, 60, 850, 60)
    With myShape.TextFrame.TextRange
        .Text = "SOLUSI: Arsitektur Otomasi AI Event-Driven 24/7"
        .Font.Name = "Calibri"
        .Font.Size = 28
        .Font.Bold = msoTrue
        .Font.Color.RGB = RGB(16, 185, 129)
    End With
    
    Set myShape = mySlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 80, 150, 850, 350)
    With myShape.TextFrame.TextRange
        .Text = "• Webhook Trigger: Menangkap interaksi WhatsApp & Form masuk dalam 50 milidetik." & vbCrLf & vbCrLf & _
                "• AI Agent Parsing: Mengekstrak niat beli, jadwal rapat, dan data PIC secara instan." & vbCrLf & vbCrLf & _
                "• Auto-Sync Ekosistem: Menghubungkan Google Calendar, Google Sheets, dan CRM tanpa campur tangan manusia."
        .Font.Name = "Calibri"
        .Font.Size = 20
        .Font.Color.RGB = RGB(51, 65, 85)
    End With

    ' SLIDE 4: ROI & IMPACT
    Set mySlide = pptPres.Slides.Add(4, ppLayoutCustom)
    mySlide.Background.Fill.Solid
    mySlide.Background.Fill.ForeColor.RGB = RGB(15, 23, 42) ' Dark Navy
    
    Set myShape = mySlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 80, 60, 850, 60)
    With myShape.TextFrame.TextRange
        .Text = "IMPACT: Estimasi Penghematan Rp 15.000.000 / Bulan"
        .Font.Name = "Calibri"
        .Font.Size = 28
        .Font.Bold = msoTrue
        .Font.Color.RGB = RGB(250, 204, 21)
    End With
    
    Set myShape = mySlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 80, 160, 850, 320)
    With myShape.TextFrame.TextRange
        .Text = "1. Efisiensi Jam Kerja: Menghemat 120 jam operasional tim per bulan." & vbCrLf & vbCrLf & _
                "2. Kecepatan Konversi: Tingkat closing naik 35% karena respon instan dalam 2 detik." & vbCrLf & vbCrLf & _
                "3. Skalabilitas: Mampu menangani 10.000 pesan tanpa perlu menambah staf admin tambahan."
        .Font.Name = "Calibri"
        .Font.Size = 20
        .Font.Color.RGB = RGB(241, 245, 249)
    End With

    MsgBox "Sukses! 4 Slide Presentasi Konsultan Berhasil Dibuat Otomatis!", vbInformation, "Pangestu AI Presentation Architect"
End Sub

