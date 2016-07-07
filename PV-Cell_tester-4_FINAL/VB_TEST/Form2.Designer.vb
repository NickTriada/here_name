<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_stat
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form_stat))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Lbl_stat = New System.Windows.Forms.Label()
        Me.Lbl_stat2 = New System.Windows.Forms.Label()
        Me.Chart = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.Rich_Lbl_stat = New System.Windows.Forms.RichTextBox()
        CType(Me.Chart, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(147, -1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Statistics"
        '
        'Lbl_stat
        '
        Me.Lbl_stat.AutoSize = True
        Me.Lbl_stat.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_stat.Location = New System.Drawing.Point(198, 29)
        Me.Lbl_stat.Name = "Lbl_stat"
        Me.Lbl_stat.Size = New System.Drawing.Size(0, 15)
        Me.Lbl_stat.TabIndex = 1
        '
        'Lbl_stat2
        '
        Me.Lbl_stat2.AutoSize = True
        Me.Lbl_stat2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_stat2.Location = New System.Drawing.Point(14, 30)
        Me.Lbl_stat2.Name = "Lbl_stat2"
        Me.Lbl_stat2.Size = New System.Drawing.Size(0, 15)
        Me.Lbl_stat2.TabIndex = 2
        '
        'Chart
        '
        ChartArea1.Name = "ChartArea1"
        Me.Chart.ChartAreas.Add(ChartArea1)
        Legend1.Name = "Legend1"
        Me.Chart.Legends.Add(Legend1)
        Me.Chart.Location = New System.Drawing.Point(1, 377)
        Me.Chart.Name = "Chart"
        Me.Chart.Size = New System.Drawing.Size(396, 274)
        Me.Chart.TabIndex = 3
        Me.Chart.Text = "Chart"
        '
        'Rich_Lbl_stat
        '
        Me.Rich_Lbl_stat.Font = New System.Drawing.Font("Microsoft Sans Serif", 5.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Rich_Lbl_stat.Location = New System.Drawing.Point(1, 653)
        Me.Rich_Lbl_stat.Name = "Rich_Lbl_stat"
        Me.Rich_Lbl_stat.Size = New System.Drawing.Size(396, 20)
        Me.Rich_Lbl_stat.TabIndex = 4
        Me.Rich_Lbl_stat.Text = ""
        '
        'Form_stat
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(399, 674)
        Me.Controls.Add(Me.Rich_Lbl_stat)
        Me.Controls.Add(Me.Chart)
        Me.Controls.Add(Me.Lbl_stat2)
        Me.Controls.Add(Me.Lbl_stat)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form_stat"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Statistics"
        CType(Me.Chart, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Lbl_stat As Label
    Friend WithEvents Lbl_stat2 As Label
    Friend WithEvents Chart As DataVisualization.Charting.Chart
    Friend WithEvents Rich_Lbl_stat As RichTextBox
End Class
