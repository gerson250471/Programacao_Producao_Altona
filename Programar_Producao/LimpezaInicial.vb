Private Sub LimpezaInicial()

    P06.AutoFilterMode = False                          'Rel 01
    P07.AutoFilterMode = False                          'Rel 02
    P08.AutoFilterMode = False                          'Rel 03
    P09.AutoFilterMode = False                          'Rel 04
    P10.AutoFilterMode = False                          'Rel 05
    P16.AutoFilterMode = False                          'Rel 07
    P18.AutoFilterMode = False                          'Rel 08
    P03.Range("A2:EZ70000") = ""                        'Temp 00
    P05.Range("A2:EZ70000") = ""                        'Temp 02
    P06.Range("A5:Z30000").EntireRow.Delete             'Rel 01
    P07.Range("A5:Z30000").EntireRow.Delete             'Rel 02
    P08.Range("A5:Z30000").EntireRow.Delete             'Rel 03
    P09.Range("A5:Z30000").EntireRow.Delete             'Rel 04
    P10.Range("A5:Z30000").EntireRow.Delete             'Rel 05
    P16.Range("A5:Z30000").EntireRow.Delete             'Rel 07
    P18.Range("A5:Z30000").EntireRow.Delete             'Rel 08
    
End Sub