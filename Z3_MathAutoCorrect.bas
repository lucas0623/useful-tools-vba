Attribute VB_Name = "Z3_MathAutoCorrect"
Sub AutoCorrectMathSymbols()

    ' Kinen Ma
    ' 2024-11-07
    ' Version 1.0
    ' Similar to Microsoft word math autocorrection function, run sub to math autocorrect on selected cells

    Dim cell As Range
    Dim text As String
    Dim formulaParts() As String
    Dim inQuotes As Boolean
    Dim i As Integer
    Dim corrections As Object
    Dim key As Variant
    
    ' Create a dictionary of LaTeX-like commands and their corresponding Unicode symbols
    Set corrections = CreateObject("Scripting.Dictionary")
    
    ' Lowercase Greek letters
    corrections.Add "\alpha", ChrW(&H3B1)   ' a (alpha)
    corrections.Add "\beta", ChrW(&H3B2)    ' ß (beta)
    corrections.Add "\gamma", ChrW(&H3B3)   ' ? (gamma)
    corrections.Add "\delta", ChrW(&H3B4)   ' d (delta)
    corrections.Add "\epsilon", ChrW(&H3B5) ' e (epsilon)
    corrections.Add "\zeta", ChrW(&H3B6)    ' ? (zeta)
    corrections.Add "\eta", ChrW(&H3B7)     ' ? (eta)
    corrections.Add "\theta", ChrW(&H3B8)   ' ? (theta)
    corrections.Add "\iota", ChrW(&H3B9)    ' ? (iota)
    corrections.Add "\kappa", ChrW(&H3BA)   ' ? (kappa)
    corrections.Add "\lambda", ChrW(&H3BB)  ' ? (lambda)
    corrections.Add "\mu", ChrW(&H3BC)      ' µ (mu)
    corrections.Add "\nu", ChrW(&H3BD)      ' ? (nu)
    corrections.Add "\xi", ChrW(&H3BE)      ' ? (xi)
    corrections.Add "\omicron", ChrW(&H3BF) ' ? (omicron)
    corrections.Add "\pi", ChrW(&H3C0)      ' p (pi)
    corrections.Add "\rho", ChrW(&H3C1)     ' ? (rho)
    corrections.Add "\sigma", ChrW(&H3C3)   ' s (sigma)
    corrections.Add "\tau", ChrW(&H3C4)     ' t (tau)
    corrections.Add "\upsilon", ChrW(&H3C5) ' ? (upsilon)
    corrections.Add "\phi", ChrW(&H3C6)     ' f (phi)
    corrections.Add "\chi", ChrW(&H3C7)     ' ? (chi)
    corrections.Add "\psi", ChrW(&H3C8)     ' ? (psi)
    corrections.Add "\omega", ChrW(&H3C9)   ' ? (omega)
    
    ' Uppercase Greek letters
    corrections.Add "\Alpha", ChrW(&H391)   ' ? (Alpha)
    corrections.Add "\Beta", ChrW(&H392)    ' ? (Beta)
    corrections.Add "\Gamma", ChrW(&H393)   ' G (Gamma)
    corrections.Add "\Delta", ChrW(&H394)   ' ? (Delta)
    corrections.Add "\Epsilon", ChrW(&H395) ' ? (Epsilon)
    corrections.Add "\Zeta", ChrW(&H396)    ' ? (Zeta)
    corrections.Add "\Eta", ChrW(&H397)     ' ? (Eta)
    corrections.Add "\Theta", ChrW(&H398)   ' T (Theta)
    corrections.Add "\Iota", ChrW(&H399)    ' ? (Iota)
    corrections.Add "\Kappa", ChrW(&H39A)   ' ? (Kappa)
    corrections.Add "\Lambda", ChrW(&H39B)  ' ? (Lambda)
    corrections.Add "\Mu", ChrW(&H39C)      ' ? (Mu)
    corrections.Add "\Nu", ChrW(&H39D)      ' ? (Nu)
    corrections.Add "\Xi", ChrW(&H39E)      ' ? (Xi)
    corrections.Add "\Omicron", ChrW(&H39F) ' ? (Omicron)
    corrections.Add "\Pi", ChrW(&H3A0)      ' ? (Pi)
    corrections.Add "\Rho", ChrW(&H3A1)     ' ? (Rho)
    corrections.Add "\Sigma", ChrW(&H3A3)   ' S (Sigma)
    corrections.Add "\Tau", ChrW(&H3A4)     ' ? (Tau)
    corrections.Add "\Upsilon", ChrW(&H3A5) ' ? (Upsilon)
    corrections.Add "\Phi", ChrW(&H3A6)     ' F (Phi)
    corrections.Add "\Chi", ChrW(&H3A7)     ' ? (Chi)
    corrections.Add "\Psi", ChrW(&H3A8)     ' ? (Psi)
    corrections.Add "\Omega", ChrW(&H3A9)   ' O (Omega)
    
    ' Common mathematical symbols
    corrections.Add "\infinity", ChrW(&H221E)      ' 8 (Infinity, limitless quantity)
    corrections.Add "\int", ChrW(&H222B)           ' ? (Integral, calculus operator)
    corrections.Add "\sum", ChrW(&H2211)           ' ? (Summation, summing a series)
    corrections.Add "\prod", ChrW(&H220F)          ' ? (Product, product of a sequence)
    corrections.Add "\partial", ChrW(&H2202)       ' ? (Partial derivative)
    corrections.Add "\approx", ChrW(&H2248)        ' ˜ (Approximately equal)
    corrections.Add "\neq", ChrW(&H2260)           ' ? (Not equal to)
    corrections.Add "\leq", ChrW(&H2264)           ' = (Less than or equal to)
    corrections.Add "\geq", ChrW(&H2265)           ' = (Greater than or equal to)
    corrections.Add "\pm", ChrW(&HB1)              ' ± (Plus-minus, indicates uncertainty)
    corrections.Add "\times", ChrW(&HD7)           ' × (Multiplication sign)
    corrections.Add "\div", ChrW(&HF7)             ' ÷ (Division sign)
    corrections.Add "\sqrt", ChrW(&H221A)          ' v (Square root)
    corrections.Add "\in", ChrW(&H2208)            ' ? (Element of a set)
    corrections.Add "\notin", ChrW(&H2209)         ' ? (Not an element of a set)
    corrections.Add "\subset", ChrW(&H2282)        ' ? (Subset)
    corrections.Add "\supset", ChrW(&H2283)        ' ? (Superset)
    corrections.Add "\subseteq", ChrW(&H2286)      ' ? (Subset or equal)
    corrections.Add "\supseteq", ChrW(&H2287)      ' ? (Superset or equal)
    corrections.Add "\cup", ChrW(&H222A)           ' ? (Union of sets)
    corrections.Add "\cap", ChrW(&H2229)           ' n (Intersection of sets)
    corrections.Add "\and", ChrW(&H2227)           ' ? (Logical AND)
    corrections.Add "\or", ChrW(&H2228)            ' ? (Logical OR)
    corrections.Add "\forall", ChrW(&H2200)        ' ? (For all)
    corrections.Add "\exists", ChrW(&H2203)        ' ? (There exists)
    corrections.Add "\nexists", ChrW(&H2204)       ' ? (There does not exist)
    corrections.Add "\Rightarrow", ChrW(&H21D2)    ' ? (Implies, logical implication)
    corrections.Add "\Leftrightarrow", ChrW(&H21D4) ' ? (If and only if, logical equivalence)
    corrections.Add "\rightarrow", ChrW(&H2192)    ' ? (Right arrow, function mapping)
    corrections.Add "\leftarrow", ChrW(&H2190)     ' ? (Left arrow, inverse functions)
    corrections.Add "\Leftarrow", ChrW(8656)     ' ? (Left arrow, logical implication)
    corrections.Add "\leftrightarrow", ChrW(&H2194) ' ? (Bidirectional arrow, equivalence)
    corrections.Add "\==>", ChrW(&H21D2)    ' ? (Implies, logical implication)
    corrections.Add "\<==>", ChrW(&H21D4) ' ? (If and only if, logical equivalence)
    corrections.Add "\-->", ChrW(&H2192)    ' ? (Right arrow, function mapping)
    corrections.Add "\<--", ChrW(&H2190)     ' ? (Left arrow, inverse functions)
    corrections.Add "\<==", ChrW(8656)     ' ? (Left arrow, logical implication)
    corrections.Add "\<-->", ChrW(&H2194) ' ? (Bidirectional arrow, equivalence)
    corrections.Add "\angle", ChrW(&H2220)         ' ? (Angle)
    corrections.Add "\perp", ChrW(&H27C2)          ' ? (Perpendicular)
    corrections.Add "\parallel", ChrW(&H2225)      ' ? (Parallel)
    corrections.Add "\nabla", ChrW(&H2207)         ' ? (Nabla, gradient operator)
    corrections.Add "\cong", ChrW(&H2245)          ' ? (Congruent to)
    corrections.Add "\equiv", ChrW(&H2261)         ' = (Identically equal to)
    corrections.Add "\therefore", ChrW(&H2234)     ' ? (Therefore)
    corrections.Add "\because", ChrW(&H2235)       ' ? (Because)
    corrections.Add "\propto", ChrW(&H221D)        ' ? (Proportional to)
    corrections.Add "\infty", ChrW(&H221E)         ' 8 (Infinity, duplicate for convenience)
    corrections.Add "\aleph", ChrW(&H2135)         ' ? (Aleph, cardinality of infinite sets)
    corrections.Add "\implies", ChrW(&H21D2)       ' ? (Implies, duplicate for convenience)
    corrections.Add "\iff", ChrW(&H21D4)           ' ? (If and only if, duplicate for convenience)
    corrections.Add "\bot", ChrW(&H22A5)           ' ? (Bottom, contradiction)
    corrections.Add "\top", ChrW(&H22A4)           ' ? (Top, tautology)
    corrections.Add "\vdash", ChrW(&H22A2)         ' ? (Provable, syntactic consequence)
    corrections.Add "\dashv", ChrW(&H22A3)         ' ? (Right tack, semantic consequence)
    corrections.Add "\vdots", ChrW(&H22EE)         ' ? (Vertical ellipsis)
    corrections.Add "\cdots", ChrW(&H22EF)         ' ? (Center ellipsis)
    corrections.Add "\ldots", ChrW(&H2026)         ' … (Horizontal ellipsis)
    corrections.Add "\<=", ChrW(&H2264)             ' = (Less than or equal to)
    corrections.Add "\>=", ChrW(&H2265)             ' = (Greater than or equal to)
    corrections.Add "\==", ChrW(&H2261)             ' = (Identical to)
    corrections.Add "\!=", ChrW(&H2260)             ' ? (Not equal to)
    corrections.Add "\~=", ChrW(&H2248)             ' ˜ (Approximately equal)
    corrections.Add "\plusminus", ChrW(&HB1)       ' ± (Plus-minus sign)
    corrections.Add "\minus", ChrW(&H2212)         ' - (Minus sign)
    corrections.Add "\cdot", ChrW(&H22C5)          ' · (Dot operator)
    corrections.Add "\ast", ChrW(&H2217)           ' * (Asterisk operator)
    corrections.Add "\star", ChrW(&H22C6)          ' ? (Star operator)
    corrections.Add "\circ", ChrW(&H2218)          ' ° (Ring operator)
    corrections.Add "\bullet", ChrW(&H2022)        ' • (Bullet operator)
    corrections.Add "\divides", ChrW(&H2223)       ' | (Divides)
    corrections.Add "\nmid", ChrW(&H2224)          ' ? (Does not divide)
    corrections.Add "\nparallel", ChrW(&H2226)     ' ? (Not parallel to)
    corrections.Add "\measuredangle", ChrW(&H2221) ' ? (Measured angle)
    corrections.Add "\sphericalangle", ChrW(&H2222) ' ? (Spherical angle)
    corrections.Add "\notperp", ChrW(&H22A5)       ' ? (Not perpendicular)
    corrections.Add "\wedge", ChrW(&H2227)         ' ? (Logical AND)
    corrections.Add "\vee", ChrW(&H2228)           ' ? (Logical OR)
    corrections.Add "\neg", ChrW(&HAC)             ' ¬ (Logical NOT)
    corrections.Add "\oplus", ChrW(&H2295)         ' ? (Circled plus)
    corrections.Add "\ominus", ChrW(&H2296)        ' ? (Circled minus)
    corrections.Add "\otimes", ChrW(&H2297)        ' ? (Circled times)
    corrections.Add "\oslash", ChrW(&H2298)        ' ? (Circled division slash)
    corrections.Add "\odot", ChrW(&H2299)          ' ? (Circled dot operator)
    corrections.Add "\union", ChrW(&H222A)         ' ? (Union)
    corrections.Add "\intersection", ChrW(&H2229)  ' n (Intersection)
    corrections.Add "\uplus", ChrW(&H228E)         ' ? (Multiset union)
    corrections.Add "\setminus", ChrW(&H2216)      ' \ (Set minus)
    corrections.Add "\complement", ChrW(&H2201)    ' ? (Complement)
    corrections.Add "\hbar", ChrW(&H210F)          ' ? (Reduced Planck's constant)
    corrections.Add "\proportional", ChrW(&H221D)  ' ? (Proportional to)
    corrections.Add "\varnothing", ChrW(&H2205)    ' Ø (Empty set)
    corrections.Add "\triangle", ChrW(&H25B3)      ' ? (Triangle)
    corrections.Add "\triangleq", ChrW(&H225C)     ' ? (Equal by definition)
    corrections.Add "\dot", ChrW(&H22C5)           ' · (Dot operator)
    corrections.Add "\ddot", ChrW(&H308)           ' ¨ (Double dot, used in calculus)
    corrections.Add "\cancel", ChrW(&H336)         ' ? (Strike through, used to cancel terms)
    corrections.Add "\square", ChrW(&H25A1)        ' ? (Square, often used in proofs)
    corrections.Add "\blacksquare", ChrW(&H25A0)   ' ¦ (Black square, end of proof)
    corrections.Add "\checkmark", ChrW(&H2713)     ' ? (Check mark)
    corrections.Add "\dagger", ChrW(&H2020)        ' † (Dagger, used in matrices)
    corrections.Add "\ddagger", ChrW(&H2021)       ' ‡ (Double dagger, used in matrices)
    corrections.Add "\backslash", ChrW(&H2216)     ' \ (Set difference)
    corrections.Add "\Re", ChrW(&H211C)            ' R (Real part of complex number)
    corrections.Add "\Im", ChrW(&H2111)            ' I (Imaginary part of complex number)
    corrections.Add "\wp", ChrW(&H2118)            ' P (Weierstrass p function)

    ' Superscript numbers
    corrections.Add "\^0", ChrW(&H2070)            ' ° (Superscript Zero)
    corrections.Add "\^1", ChrW(&HB9)              ' ¹ (Superscript One)
    corrections.Add "\^2", ChrW(&HB2)              ' ² (Superscript Two)
    corrections.Add "\^3", ChrW(&HB3)              ' ³ (Superscript Three)
    corrections.Add "\^4", ChrW(&H2074)            ' 4 (Superscript Four)
    corrections.Add "\^5", ChrW(&H2075)            ' 5 (Superscript Five)
    corrections.Add "\^6", ChrW(&H2076)            ' 6 (Superscript Six)
    corrections.Add "\^7", ChrW(&H2077)            ' 7 (Superscript Seven)
    corrections.Add "\^8", ChrW(&H2078)            ' 8 (Superscript Eight)
    corrections.Add "\^9", ChrW(&H2079)            ' ? (Superscript Nine)
    
    ' Subscript numbers
    corrections.Add "\~^0", ChrW(&H2080)            ' 0 (Subscript Zero)
    corrections.Add "\~^1", ChrW(&H2081)            ' 1 (Subscript One)
    corrections.Add "\~^2", ChrW(&H2082)            ' 2 (Subscript Two)
    corrections.Add "\~^3", ChrW(&H2083)            ' 3 (Subscript Three)
    corrections.Add "\~^4", ChrW(&H2084)            ' 4 (Subscript Four)
    corrections.Add "\~^5", ChrW(&H2085)            ' 5 (Subscript Five)
    corrections.Add "\~^6", ChrW(&H2086)            ' 6 (Subscript Six)
    corrections.Add "\~^7", ChrW(&H2087)            ' 7 (Subscript Seven)
    corrections.Add "\~^8", ChrW(&H2088)            ' 8 (Subscript Eight)
    corrections.Add "\~^9", ChrW(&H2089)            ' 9 (Subscript Nine)




    ' Loop through each cell in the selection
    For Each cell In Selection
        If cell.HasFormula Then
            ' If the cell contains a formula, process only the parts within quotes
            text = cell.formula
            formulaParts = Split(text, """")
            inQuotes = False
            
            For i = LBound(formulaParts) To UBound(formulaParts)
                If inQuotes Then
                    ' Replace text within quotes
                    For Each key In corrections.Keys
                        formulaParts(i) = Replace(formulaParts(i), key, corrections(key))
                    Next key
                End If
                inQuotes = Not inQuotes
            Next i
            
            ' Reconstruct the formula with corrected text
            cell.formula = Join(formulaParts, """")
            
        Else
            ' If the cell contains plain text, process the entire cell
            text = cell.Value
            For Each key In corrections.Keys
                text = Replace(text, key, corrections(key))
            Next key
            cell.Value = text
        End If
    Next cell
End Sub
