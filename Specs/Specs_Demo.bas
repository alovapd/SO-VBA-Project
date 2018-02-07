Attribute VB_Name = "Specs_Demo"
Public Function spec() As SpecSuite

    Set spec = New SpecSuite
    spec.Description = "DemoSpecs"
    
    With spec.It("should equal 2")
        .Expect(1 + 5).ToEqual 2
    End With
    
End Function
