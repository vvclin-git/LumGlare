Imports System.Math
Module MatMath
    Function eulerXYZ(alpha As Double, beta As Double, gamma As Double, input As Pt3D) As Pt3D
        Dim x0, y0, z0 As Double
        Dim x1, y1, z1 As Double
        x0 = input.getX()
        y0 = input.getY()
        z0 = input.getZ()
        'transform degrees to radians
        alpha = alpha * Math.PI / 180
        beta = beta * Math.PI / 180
        gamma = gamma * Math.PI / 180
        x1 = Cos(beta) * Cos(gamma) * x0 - Cos(beta) * Sin(gamma) * y0 + Sin(beta) * z0
        y1 = (Cos(alpha) * Sin(gamma) + Cos(gamma) * Sin(alpha) * Sin(beta)) * x0 _
                        + (Cos(alpha) * Cos(gamma) - Sin(alpha) * Sin(beta) * Sin(gamma)) * y0 _
                        - Cos(beta) * Sin(alpha) * z0
        z1 = (Sin(alpha) * Sin(gamma) - Cos(alpha) * Cos(gamma) * Sin(beta)) * x0 _
                        + (Cos(gamma) * Sin(alpha) + Cos(alpha) * Sin(beta) * Sin(gamma)) * y0 _
                        + Cos(alpha) * Cos(beta) * z0
        eulerXYZ = New Pt3D(x1, y1, z1)
    End Function
    Function ltTransXYZ(alpha As Double, beta As Double, gamma As Double, trans As Pt3D, input As Pt3D) As Pt3D
        Dim ltRotated As Pt3D
        ltRotated = eulerXYZ(-alpha, -beta, gamma, input)
        ltTransXYZ = New Pt3D(ltRotated.getX + trans.getX, ltRotated.getY + trans.getY, ltRotated.getZ + trans.getZ)
    End Function
    Class Pt3D
        Dim x As Double
        Dim y As Double
        Dim z As Double
        Public Sub New(x As Double, y As Double, z As Double)
            Me.x = x
            Me.y = y
            Me.z = z
        End Sub
        Public Function getX() As Double
            getX = Me.x
        End Function
        Public Function getY() As Double
            getY = Me.y
        End Function
        Public Function getZ() As Double
            getZ = Me.z
        End Function
        Public Overrides Function ToString() As String
            Return "(" + Str(x) + ", " + Str(y) + ", " + Str(z) + ")"
        End Function
        Public Function toLTXYZ() As String
            Return "XYZ " + Str(x).Trim + "," + Str(y).Trim + "," + Str(z).Trim
        End Function
    End Class

End Module
