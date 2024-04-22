VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   3480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   4
      Top             =   3360
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Training starten"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   3360
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   10575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   4800
      Width           =   8295
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   10815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Anzahl Trainings-Beispiele pro Trainings-Runde:
Const gAmountOfExamplesPerTraining = 1000

'Die Maximal-Werte für Quadratmeter und Preis:
Const cMaxValueM2 = 320
Const cMaxValuePrice = 1500000

' Status of NeuralNetworkInitialized
Dim blnNeuralNetworkInitialized As Boolean
Dim blnSampleRunning            As Boolean

' Statische Variablen für die Netzwerkparameter
Dim weights1() As Double
Dim weights2() As Double
Dim bias1() As Double
Dim bias2() As Double

' Statische Variable für die Zwischenergebnisse
Dim layer1(1 To 2) As Double
Dim output(1 To 3) As Double

'Zählwerte
Dim lAmountTrainedExamples As Long
Dim lAmountPositiveSamples As Long
Dim lAmountNegativeSamples As Long
Dim mdblSuccessRate        As Double

'Wenn unser Neuronales Netzwerk schon das grunsätzliche gelernt hat,
'dann macht es Sinn, das Wissen nur noch zu kalibrieren und nichts
'Neues mehr zu lernen, da wir sonst in unserem neuronalen Netz
'zu Verwirrung sorgen könnten.
'
'Am Anfang sind wir aber in der aktiven Lernphase, und da wird die
'Backpropagation solange durchgeführt, bis das neuronale Netzwerk
'die richtige Antwort sagen kann.
Dim gblnCalibrate           As Boolean


Function GetValueFromModel(m2 As Double, price As Double) As Double
'Gibt einen Wert zurück unseres Deep Learning Modells.
'Alle Parameter des Modells müssen angegeben werden.
'Dann gibt unser virtuelles "Gehirn" seine Einschätzung
'zurück.
'-------------------------------------------------------------
Dim input1 As Double, input2 As Double

  'Normalisiere die Input-Daten:
  input1 = NormalizeInput(m2, 0, cMaxValueM2) 'Max. 12500 Quadratmeter
  input2 = NormalizeInput(price, 0, cMaxValuePrice) 'Max. 990 Millionen

  ForwardPass input1, input2
  GetValueFromModel = output(3)

End Function

Sub InitializeNeuralNetwork()
    ' Netzwerkparameter initialisieren
    If blnNeuralNetworkInitialized Then Exit Sub
    
    Dim numInputs As Integer
    numInputs = 2
    
    ReDim weights1(2, numInputs)
    ReDim weights2(numInputs + 1)
    ReDim bias1(numInputs)
    ReDim bias2(1)
    
    ' Zufällige Initialisierung der Gewichte und Bias
    'Die zufällige Initialisierung der Gewichte und Schwellenwerte (Bias) ist eine gängige
    'Praxis beim Trainieren von neuronalen Netzwerken. Die Idee dahinter ist, dass eine
    'zufällige Initialisierung den Netzwerkparametern einen Ausgangspunkt gibt, von dem aus
    'sie während des Trainings optimiert werden können.
    '
    'Beim Training eines neuronalen Netzwerks geht es darum, die Gewichte und Schwellenwerte
    'anzupassen, um das Netzwerk dazu zu bringen, genaue Vorhersagen zu treffen. Indem wir
    'die Gewichte und Schwellenwerte zufällig initialisieren, stellen wir sicher, dass das
    'Netzwerk nicht von Anfang an voreingenommen oder in einem ungünstigen Zustand ist.
    'Stattdessen erhält das Netzwerk die Möglichkeit, während des Trainings die Gewichte
    'anzupassen, um die besten Ergebnisse zu erzielen.
    '
    'Die zufällige Initialisierung sorgt auch dafür, dass verschiedene Instanzen desselben
    'neuronalen Netzwerks unterschiedliche Anfangszustände haben. Das kann hilfreich sein,
    'um sicherzustellen, dass das Netzwerk unterschiedliche Muster in den Daten erlernen kann
    'und nicht in einem lokalen Optimum stecken bleibt.
    '
    'Es ist wichtig zu beachten, dass die zufällige Initialisierung der Gewichte und
    'Schwellenwerte keinen direkten Einfluss auf die Funktion oder den Erfolg des neuronalen
    'Netzwerks hat. Sie dient lediglich dazu, dem Netzwerk einen Startpunkt zu geben, von dem
    'aus es seine Gewichte optimieren kann.
    '---------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    For i = 1 To numInputs
        For j = 1 To 2
            weights1(j, i) = Rnd()
        Next j
        bias1(i) = Rnd()
    Next i
    
    For i = 1 To numInputs + 1
        weights2(i) = Rnd()
    Next i
    
    bias2(1) = Rnd()
    
    blnNeuralNetworkInitialized = True
End Sub

Function Activate(ByVal x As Double) As Double
'Aktivierungsfunktion. Diese Funktion aktiviert ein Neuron aufgrund eines Wertes und gibt
'diesem Node damit quasi einen inneren Status zwischen 0 und 1.
'
'Durch diese Funktion wird der normalisierte Eingabewert auf eine "aktivierte" Ausgabe
'umgewandelt, die die Aktivierung des Neurons repräsentiert. Die Ausgabe liegt immer
'zwischen 0 und 1 und kann als Mass für die Aktivierung oder Wahrscheinlichkeit
'interpretiert werden.
'
'Wir verwenden dafür die Sigmoid-Funktion.
'
'-----------------------------------------------------------------------------------------
  Activate = 1 / (1 + Exp(-x))
End Function

Sub ForwardPass(ByVal input1 As Double, ByVal input2 As Double)
'Vorwärtsberechnung der Ausgabe mit den Eingabeparametern durch das neuronale
'Netzwerk.
'
'Merke:
'
'Hier werden die Eingabeparameter (vergleichbar mit den "Eindrücken" des Menschen)
'quasi durch das virtuelle "Hirn" (=neuronales Netzwerk) geleitet.
'
'In unserem Beispiel, sind die Eigabeparameter
'
'input1     Quadratmeter der Immobilie
'input2     Preis der Immobilie
'
'Es ist empfehlenswert, die Eingabeparameter vor der Übergabe an
'diese Funktion zu normalisieren. Verwenden Sie dazu die
'NormalizeInput-Funktion.
'------------------------------------------------------------------------
    
    ' Eingabewerte setzen
    layer1(1) = input1
    layer1(2) = input2
    
    'Berechnung der Ausgabe der versteckten Schicht
    '
    ' Was passiert hier? Und für was braucht es hier versteckte Schichten?
    '
    ' In einem neuronalen Netzwerk gibt es normalerweise eine oder mehrere versteckte Schichten
    ' zwischen der Eingabeschicht und der Ausgabeschicht. Diese versteckten Schichten
    ' enthalten Neuronen, die Informationen verarbeiten und transformieren, um die
    ' endgültige Ausgabe des Netzwerks zu erzeugen.
    '
    ' Die Berechnung der Ausgabe der versteckten Schicht beinhaltet die Anwendung einer
    ' Aktivierungsfunktion auf eine Kombination der Eingabewerte und der Gewichte der
    ' versteckten Schicht. Die Aktivierungsfunktion dient dazu, die Aktivierung jedes
    ' Neurons in der versteckten Schicht zu bestimmen.
    '
    ' Im gegebenen Beispielcode wird die Sigmoid-Aktivierungsfunktion verwendet. Die
    ' Aktivierung jedes Neurons in der versteckten Schicht wird berechnet, indem die
    ' gewichteten Eingabewerte (Multiplikation der Gewichte mit den Eingabewerten) und
    ' die Schwellenwerte (Bias) summiert und dann auf die Aktivierungsfunktion angewendet
    ' werden.
    '
    ' Dieser Schritt wird für jedes Neuron in der versteckten Schicht wiederholt,
    ' um die Ausgaben der versteckten Schicht zu berechnen, die dann als Eingabe für
    ' die Ausgabeschicht dienen.
    '
    '-----------------------------------------------------------------------------------------
    output(1) = Activate(weights1(1, 1) * input1 + weights1(1, 2) * input2 + bias1(1))
    output(2) = Activate(weights1(2, 1) * input1 + weights1(2, 2) * input2 + bias1(2))
    
    'Tipp:
    'Wenn die Eingabewerte input1 und input2 grössere Zahlen sind, können die gewichteten
    'Summen
    '
    '   weights1(1, 1) * input1 + weights1(1, 2) * input2 + bias1(1)
    ' und
    '   weights1(2, 1) * input1 + weights1(2, 2) * input2 + bias1(2)
    '
    'sehr gross werden. Wenn die Eingabewerte sehr gross sind, neigt die
    'Sigmoid-Aktivierungsfunktion dazu, nahe an 1 zu sättigen, was bedeutet,
    'dass die Ausgabe des Neurons fast 1 wird.
    '
    'Dies liegt daran, dass die Ableitung der Sigmoid-Funktion in der Nähe von 1
    'sehr klein ist, und dies kann dazu führen, dass die Gradienten in den Schichten,
    'die der Ausgabeschicht vorausgehen, sehr klein werden. Wenn die Gradienten zu
    'klein werden, kann dies das Lernen des Netzwerks verlangsamen oder sogar verhindern.
    '
    'Es ist wichtig zu beachten, dass diese Sättigung der Sigmoid-Funktion ein Problem
    'sein kann und als "Vanishing Gradient Problem" bekannt ist. Um dieses Problem zu
    'lösen, können andere Aktivierungsfunktionen wie die ReLU-Funktion verwendet werden,
    'die nicht von Sättigung betroffen ist oder andere Techniken wie die Initialisierung
    'der Gewichte und der Lernrate verwendet werden.
    '
    'Weil wir wollen, dass die Ausgaben nicht ausschliesslich 1 sind bei grösseren
    'Input-Werten, dann könnten wir eine Aktivierungsfunktion verwenden, die besser
    'mit großen Eingabewerten umgehen kann oder die Eingabewerte normalisieren kann,
    'um die Auswirkungen der großen Werte auf die Ausgaben zu reduzieren.
    '
    'Es ist wichtig zu beachten, dass unsere Aktivierungsfunktion die Werte deshalb
    'normalisiert und den Sigmoid auf die normalisierten Werte zurückgibt.
    
    ' Berechnung der Ausgabe
    output(3) = Activate(weights2(1) * output(1) + weights2(2) * output(2) + bias2(1))

End Sub

Sub Backpropagation(ByVal targetOutput As Double, ByVal learningRate As Double)
' Rückpropagierung der Gradienten und Aktualisierung der Netzwerkparameter
'
' Dies lernt dem Modell, dass der letzte Wert targetOutput hätte sein sollen.
'
' Hint:
'
' The learning rate is a hyperparameter -- a factor that defines the system or
' set conditions for its operation prior to the learning process -- that controls
' how much change the model experiences in response to the estimated error every
' time the model weights are altered. Learning rates that are too high may result
' in unstable training processes or the learning of a suboptimal set of weights.
' Learning rates that are too small may produce a lengthy training process that has
' the potential to get stuck.
'
' The learning rate decay method -- also called learning rate annealing or adaptive
' learning rates -- is the process of adapting the learning rate to increase
' performance and reduce training time. The easiest and most common adaptations
' of learning rate during training include techniques to reduce the learning rate
' over time.
'
'--------------------------------------------------------------------------------------
    
    ' Berechnung des Fehlers
    Dim outputError As Double
    outputError = targetOutput - output(3)
    
    ' Berechnung der Gradienten für die Ausgabeschicht
    Dim outputGradient As Double
    outputGradient = output(3) * (1 - output(3)) * outputError
    
    ' Aktualisierung der Gewichte und Bias zwischen Ausgabeschicht und versteckter Schicht
    Dim delta1 As Double, delta2 As Double
    
    delta1 = learningRate * outputGradient * output(1)
    delta2 = learningRate * outputGradient * output(2)
    
    weights2(1) = weights2(1) + delta1
    weights2(2) = weights2(2) + delta2
    
    ' Aktualisierung des Schwellenwerts der Ausgabeschicht
    bias2(1) = bias2(1) + learningRate * outputGradient
    
    ' Berechnung der Gradienten für die versteckte Schicht
    Dim hiddenGradient1 As Double, hiddenGradient2 As Double
    
    hiddenGradient1 = output(1) * (1 - output(1)) * outputGradient * weights2(1)
    hiddenGradient2 = output(2) * (1 - output(2)) * outputGradient * weights2(2)
    
    ' Aktualisierung der Gewichte und Bias zwischen Eingabeschicht und versteckter Schicht
    Dim i As Integer
    
    For i = 1 To 2
        weights1(1, i) = weights1(1, i) + learningRate * hiddenGradient1 * layer1(i)
        weights1(2, i) = weights1(2, i) + learningRate * hiddenGradient2 * layer1(i)
    Next i
    
    ' Aktualisierung der Schwellenwerte der versteckten Schicht
    bias1(1) = bias1(1) + learningRate * hiddenGradient1
    bias1(2) = bias1(2) + learningRate * hiddenGradient2

End Sub

Sub NeuralNetworkExample()
Dim learningRate As Double

    ' Beispiel für die Verwendung des neuronalen Netzwerks
    blnSampleRunning = True
    
    'Beispieldatensatz erstellen...
    GenerateSampleData
    
    ' Initialisierung des Netzwerks
    InitializeNeuralNetwork
    
    'Merke:
    'In unserem neuronalen Netzwerk haben wir festgestellt,
    'dass bessere Ergebnisse erzielt werden, wenn wir das Training
    'unserer KI in zwei Hauptphasen aufteilen: das "Pauken" und das
    '"Kalibrieren".
    '
    'Während der "Pauken"-Phase stellt unser KI-Modell sicher,
    'dass es das Gelernte korrekt versteht und beantworten kann.
    'Wenn es Lücken gibt, wird es vor dem nächsten Lernabschnitt
    'mittels BackPropagation nachtrainiert.
    '
    '->Hier unterscheidet sich unser Modell von anderen - anstatt
    '  adaptive Lernraten zu verwenden, setzen wir hier auf diesen
    '  zweiphasigen Ansatz mit "Pauken und Kalibrieren":
    
    'Die maximale Größe der learningRate (Lernrate) hängt von verschiedenen Faktoren ab,
    'einschließlich des spezifischen Optimierungsalgorithmus, den du verwendest, und
    'der Art des Problems, das du mit dem neuronalen Netzwerk löst.

    'Die Lernrate ist ein Hyperparameter, der den Schritt bestimmt, mit dem die Gewichte
    'und Schwellenwerte während des Trainings angepasst werden. Eine zu große Lernrate
    'kann dazu führen, dass das Netzwerk instabil wird und die Gewichte übermäßig stark
    'angepasst werden, was zu schlechteren Leistungen führen kann. Auf der anderen Seite
    'kann eine zu kleine Lernrate dazu führen, dass das Netzwerk langsam konvergiert
    'und viel Zeit benötigt, um zu lernen.

    'Es ist üblich, die Lernrate in einem vernünftigen Bereich zu wählen,
    'typischerweise zwischen 0,1 und 0,0001. Allerdings kann die optimale Lernrate
    'stark von Anwendung zu Anwendung variieren. Es ist oft sinnvoll, verschiedene Lernraten
    'auszuprobieren und die Leistung des Netzwerks auf einem Validierungsdatensatz zu
    'überprüfen, um die bestmögliche Lernrate zu finden.
    
    learningRate = 0.01
    
    If mdblSuccessRate > 96 Then
      'Die Erfolgsrate für das früher gelernte liegt bei über 96%.
      '->Deshalb wechseln wir nun vom "Pauken-Modus" in den
      '  "Kalibrierungs-Modus":
      gblnCalibrate = True
      If mdblSuccessRate > 99.95 Then
        'Wir setzen die Lernrate herunter, um das bereits gelernte zu
        'schützen:
        learningRate = 0.001
      End If
    Else
      'Am Anfang wollen wir das Gelernte nicht kalibrieren, sondern richtig viel
      'neues "pauken":
      gblnCalibrate = False
    End If
    
    ' Eingabewerte setzen
    Dim input1 As Double, input2 As Double
    Dim blnTrue As Boolean, dblResult As Double
    Dim iFile As Long, m2 As Long, price As Long, goodchoice As Integer
    Dim lIndex As Long, lRight  As Long, lWrong As Long, dblOldTime As Double, dblOldTime2 As Double, uiString As String
    
    iFile = FreeFile
    lIndex = 0
    lRight = 0
    lWrong = 0
    dblOldTime = Timer - 5
    dblOldTime2 = Timer - 0.25
    Open "C:\Exampledata.txt" For Input As #iFile
    Do While Not EOF(iFile)
    
      Input #iFile, m2, price, goodchoice
      
      'Merke:
      '
      '-- Wir normalisieren hier grössere Werte, damit unsere Aktivierungs-Funktion auch mit
      '   grösseren Werten umgehen kann:
      input1 = NormalizeInput(m2, 0, cMaxValueM2) 'Max. 12500 Quadratmeter
      input2 = NormalizeInput(price, 0, cMaxValuePrice) 'Max. 990 Millionen
      
      ' Vorwärtsberechnung
      ForwardPass input1, input2
      dblResult = output(3)
      
      ' Zielwert setzen
      Dim targetOutput As Double
      targetOutput = goodchoice
      
      'Rückpropagierung der Gradienten und Aktualisierung der Netzwerkparameter...
      'Anstatt Backupropagation verwenden wir die etwas bessere Funktion "Teach":
      'Backpropagation targetOutput, learningRate
      'Merke: Wir stellen hier sicher, dass wir maximal 5 mehr negative wie positive
      '       Beispiele haben, oder maximal 5 mehr positive wie negative.
      '       So ist das Verhältnis in etwa gleich.
      If goodchoice Then
        If lAmountPositiveSamples - lAmountNegativeSamples < 5 Then
          Teach targetOutput, learningRate, CDbl(m2), CDbl(price)
          lAmountTrainedExamples = lAmountTrainedExamples + 1
          lAmountPositiveSamples = lAmountPositiveSamples + 1
        Else
          'Wir haben zu viele
          '->DropOut und nicht zum lernen benutzen.
        End If
      Else
        If lAmountNegativeSamples - lAmountPositiveSamples < 5 Then
          Teach targetOutput, learningRate, CDbl(m2), CDbl(price)
          lAmountTrainedExamples = lAmountTrainedExamples + 1
          lAmountNegativeSamples = lAmountNegativeSamples + 1
        
        Else
          'Wir haben zu viele
          '->DropOut und nicht zum lernen benutzen.
        End If
      End If
      
      If dblResult >= 0.5 Then
        'Das neuronale Netzwerk sagt: Gute Wahl
        If goodchoice Then
          uiString = Format(price, "##,##0") & ".-- für " & m2 & "m2 scheint für mich eine lohnenswerte Wohnung zu sein! Ich liege richtig. Der Quadratmeterpreis liegt bei: " & Format(price / m2, "##,##0") & ".--"
          'Debug.Print uiString
          lRight = lRight + 1
          blnTrue = True
        Else
          uiString = Format(price, "##,##0") & ".-- für " & m2 & "m2 scheint für mich eine lohnenswerte Wohnung zu sein! Ich liege falsch. Der Quadratmeterpreis liegt bei: " & Format(price / m2, "##,##0") & ".--"
          Debug.Print uiString
          lWrong = lWrong + 1
          blnTrue = False
        End If
      Else
        'Das neuronale Netzwerk sagt: Schlechte Wahl
        If goodchoice = 0 Then
          uiString = Format(price, "##,##0") & ".-- für " & m2 & "m2: Ich glaube, diese Wohnung lohnt sich wahrscheinlich nicht! Ich liege richtig! Der Quadratmeterpreis liegt bei: " & Format(price / m2, "##,##0") & ".--"
          'Debug.Print uiString
          lRight = lRight + 1
          blnTrue = True
        Else
          uiString = Format(price, "##,##0") & ".-- für " & m2 & "m2: Ich glaube, diese Wohnung lohnt sich wahrscheinlich nicht! Beim nachsehen merke ich, ich ich liege damit aber falsch. Brauche mehr Training. Der Quadratmeterpreis liegt bei: " & Format(price / m2, "##,##0") & ".--"
          Debug.Print uiString
          lWrong = lWrong + 1
          blnTrue = False
        End If
      End If
            
      'Statistik über richtig / falsch nachführen:
      If Timer - dblOldTime >= 3 Then
        'Alle fünf Sekunden ein Element ausgeben...
        Label1.Caption = uiString
        
        'Farbe setzen, je nach dem, ob ich richtig oder falsch gelegen bin...
        If blnTrue Then
          Label1.ForeColor = RGB(0, 136, 0)
        Else
          Label1.ForeColor = vbRed
        End If
        
        dblOldTime = Timer
      End If
      
      lIndex = lIndex + 1
      
      If Timer - dblOldTime >= 0.25 Or lIndex = 1 Or lIndex = gAmountOfExamplesPerTraining Then
        'Statistik nachführen
        Label2.Caption = "Status:"
        Label3.Caption = "Training und Test des neuronalen Netzwerkes. Anzahl korrekt: " & Format(100 * lRight / (lRight + lWrong), "#0.00") & "%"
        Label4.Caption = "Anzahl trainierte Beispiele: " & lAmountTrainedExamples & " / Positive Samples: " & lAmountPositiveSamples & " / Negative Samples: " & lAmountNegativeSamples & " / ERR: " & lWrong
        dblOldTime2 = Timer
      End If
      
'      If lRight + lWrong > 250 Then
'        'Setze die Statistik zurück um den laufenden Durchschnitt zurückzusetzen
'        'und das besser trainierte neuronale Netzwerk schneller darzustellen...
'        lRight = 0
'        lWrong = 0
'      End If
      
      DoEvents
    Loop
    Close #iFile

    'Die Erfolgsrate wird nun festgehalten:
    mdblSuccessRate = 100 * lRight / (lRight + lWrong)
            
    Command1.Caption = "Weitere " & gAmountOfExamplesPerTraining & " Beispieldatensätze trainieren."
    Command2.Enabled = True
    blnSampleRunning = False

End Sub

Private Sub GenerateSampleData()
  'Generiert Beispieldaten
  Dim iFile As Integer, lSample As Long, m2 As Long, price As Long, m2price As Variant, goodchoice As Integer, dblOldTime As Double, sTrainingDataUI As String
  Label1.ForeColor = vbBlack
  dblOldTime = Timer - 5
  
  iFile = FreeFile
  Open "C:\Exampledata.txt" For Output As #iFile
  
  'Apartements sollen nach m2 und Preis bewertet werden, ob es ein gutes
  'Angebot ist oder nicht.
  Randomize Timer
  For lSample = 1 To gAmountOfExamplesPerTraining
    
    m2 = Int(Rnd * cMaxValueM2) + 1
    price = Fix((Int(Rnd * cMaxValuePrice) + 1) / 1000) * 1000
    
    'In unseren Beispieldaten bestimmt der Quadratmeter-Preis, ob das
    'Apartement ein gutes Angebot ist oder nicht.
    '
    'Es ist zu bemerken, dass unsere KI später nicht mehr den Q2m-Preis hinzuziehen wird
    'um zu bestimmen, ob das ein gutes oder schlechtes Angebot ist, sondern es wird das
    'dann das neuronale Netzwerk mit diesen Beispieldaten trainiert.
    m2price = CDec(price / m2)
    If m2price >= 7500 Then
      goodchoice = 0
      sTrainingDataUI = "Beispiel: " & Format(price, "##,##0") & ".-- für " & m2 & "m2 ist ein schlechter Preis. Preis pro 2m: " & Format(Fix(m2price / 100) * 100, "##,##0") & ".--"
      'Debug.Print sTrainingDataUI
    Else
      goodchoice = 1
      sTrainingDataUI = "Beispiel: " & Format(price, "##,##0") & ".-- für " & m2 & "m2 ist ein guter Preis. Preis pro 2m: " & Format(Fix(m2price / 100) * 100, "##,##0") & ".--"
      'Debug.Print sTrainingDataUI
    End If
    
    If Timer - dblOldTime >= 3 Then
      'Alle fünf Sekunden ein Element ausgeben...
      Label1.Caption = sTrainingDataUI
      Label2.Caption = "Status:"
      Label3.Caption = "Trainingsdaten erstellen..."
      dblOldTime = Timer
    End If
    DoEvents
    
    Print #iFile, m2, price, goodchoice
  Next
  
  Close #iFile
  
End Sub

Private Sub Command1_Click()
  Command1.Enabled = False
  'Deactivate Auto Training:
  Timer1.Enabled = False
  
  'Start the Training Batch:
  NeuralNetworkExample
  
  Command1.Enabled = True
  'Autostart next Training Session (for a training over night):
  Timer1.Interval = 10000
  If Not blnSampleRunning Then
    Timer1.Enabled = True
  End If
End Sub
Private Sub Command2_Click()
  Dim m2 As String, price As String, sResult As Currency
  
  Timer1.Enabled = False
  
  Do
    price = Val(InputBox("Wie vie kostet die Immobilie? Max: " & cMaxValuePrice))
  Loop Until price <= cMaxValuePrice
  Do
    m2 = Val(InputBox("Wie gross ist die Immobilie in Quadratmeter? Max: " & cMaxValueM2))
  Loop Until m2 <= cMaxValueM2
  
  If price = 0 Or m2 = 0 Then Exit Sub
  
  sResult = GetValueFromModel(CDbl(m2), CDbl(price))
  
  If sResult >= 0.5 Then
    'Das neuronale Netzwerk sagt: Gute Wahl
    MsgBox "Ich denke, " & Format(price, "##,##0") & ".-- ist ein guter Preis für " & m2 & "m2! Was denkst Du? Der Quadratmeterpreis liegt bei: " & Format(price / m2, "##,##0") & ".--"
  Else
    'Das neuronale Netzwerk sagt: Schlechte Wahl
    MsgBox "Ich glaube nicht, dass " & Format(price, "##,##0") & ".-- ein guter Preis wäre für eine " & m2 & " Quadratmeter-Immobilie! Der Quadratmeterpreis liegt doch bei saftigen: " & Format(price / m2, "##,##0") & ".--"
  End If

  If Not blnSampleRunning Then
    Timer1.Enabled = True
  End If
  
End Sub

Private Sub Form_Load()
  mdblSuccessRate = 0
  Label1.Caption = "Herzlich willkommen in der Demo für die Erstellung eines neuronalen Netzwerkes."
  Label2.Caption = "Status:"
  Label3.Caption = "Bereit."
  Label4.Caption = "Anzahl trainierte Beispiele: " & lAmountTrainedExamples
  Label4.WordWrap = True
  Command1.Caption = "Starte das Training."
  Command2.Caption = "Neuronales Netzwerk abfragen"
  Command2.Enabled = False 'Erst trainieren, dann abfragen!
  Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
  'Start next load of Training automatically:
  Command1.Enabled = False
  Command1.Value = True
  Command1.Enabled = True
End Sub

Sub Teach(ByVal targetOutput As Double, ByVal learningRate As Double, m2 As Double, price As Double)
'Anwendungsbeispiel:
'
'm2 = 1000: price = 100000 : good = 1 : ? GetValueFromModel(cdbl(m2), cdbl(price)) : teach good,0.5,  cdbl(m2), cdbl(price)
'
'Eine vereinfachte Funktionen für die Anwendung des Deep Learning Modells:
'
'Lernt dem Modell, dass der letzte Wert targetOutput hätte sein sollen.
'
'The learning rate is a hyperparameter -- a factor that defines the system or
'set conditions for its operation prior to the learning process -- that controls
'how much change the model experiences in response to the estimated error every
'time the model weights are altered. Learning rates that are too high may result
'in unstable training processes or the learning of a suboptimal set of weights.
'Learning rates that are too small may produce a lengthy training process that has
'the potential to get stuck.
'
'The learning rate decay method -- also called learning rate annealing or adaptive
'learning rates -- is the process of adapting the learning rate to increase
'performance and reduce training time. The easiest and most common adaptations
'of learning rate during training include techniques to reduce the learning rate
'over time.
'--------------------------------------------------------------------------------------
Dim diff    As Double
Dim delta   As Double
Dim goal    As Double

Dim dblOldTime As Double

  dblOldTime = Timer - 0.1

  diff = targetOutput - output(3)
  delta = diff / 2
  goal = output(3) + delta
  
  'ula, 24.7.2023
  'Nur auf die Grenze 0.5 trainieren - für den Rest haben wir die Learningrate!
  goal = 0.5
  
  If targetOutput > 0.5 Then
    If GetValueFromModel(m2, price) < 0.5 Then
      'Training notwendig!
      '->Trainingsziel festlegen...
      diff = targetOutput - output(3)
      delta = diff / 2
      goal = output(3) + delta
    Else
      'Kein Training notwendig...
      Exit Sub
    End If
  Else
    If GetValueFromModel(m2, price) > 0.5 Then
      'Training notwendig!
      '->Trainingsziel festlegen...
      diff = targetOutput - output(3)
      delta = diff / 2
      goal = output(3) + delta
    Else
      'Kein Training notwendig...
      Exit Sub
    End If
  End If
  
  'Merke:
  'Wenn unser Neuronales Netzwerk schon das grunsätzliche gelernt hat,
  'dann macht es Sinn, das Wissen nur noch zu kalibrieren und nichts
  'Neues mehr zu lernen, da wir sonst in unserem neuronalen Netz
  'zu Verwirrung sorgen könnten.
  '
  'Am Anfang sind wir aber in der aktiven Lernphase, und da wird die
  'Backpropagation solange durchgeführt, bis das neuronale Netzwerk
  'die richtige Antwort sagen kann.
  '---
  
  If targetOutput >= 0.5 Then
    Do Until GetValueFromModel(m2, price) > goal
      Backpropagation targetOutput, learningRate
      
      If gblnCalibrate Then
        'Beim Kalibrieren stellen wir nicht sicher, dass unser neuronales
        'Netzwerk die richtige Antwort kennt, sondern machen nur einmal
        'eine Backpropagation um die Gewichtungen im Netzwerk "slightly"
        'anzupassen:
        Exit Sub
      End If
      
      If Timer - dblOldTime >= 0.1 Then
        DoEvents
        dblOldTime = Timer
      End If
    Loop
  Else
    Do Until GetValueFromModel(m2, price) < goal
      Backpropagation targetOutput, learningRate
      
      If gblnCalibrate Then
        'Beim Kalibrieren stellen wir nicht sicher, dass unser neuronales
        'Netzwerk die richtige Antwort kennt, sondern machen nur einmal
        'eine Backpropagation um die Gewichtungen im Netzwerk "slightly"
        'anzupassen:
        Exit Sub
      End If
      
      If Timer - dblOldTime >= 0.1 Then
        DoEvents
        dblOldTime = Timer
      End If
    Loop
  End If
  
End Sub

Function NormalizeInput(ByVal original_value As Double, ByVal min_value As Double, ByVal max_value As Double) As Double
    NormalizeInput = (original_value - min_value) / (max_value - min_value)
End Function
