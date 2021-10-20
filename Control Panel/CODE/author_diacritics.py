import pickle
import ftfy

special_char = {
                'Ã¢â‚¬â„¢'  :"'",   
                'Ã¢â€žâ€œ'  :'ℓ',   
                'Ã¢â€ â€œ'  :'↓',
                'Ã¢â‚¬â€œ'  :'—',   
                'Ã¢â‚¬â€“'  :'‖',         
                'Ã¢â‚¬ï¿½'  :'”' ,                            
                'Ã¢Ë†â€š'   :'∂',
                'Ã¢Ë†â€¢'   :'/',  
                'Ã¢â€°Â«'   :'>>',
                'Ã¢â‚¬Å“'   :'“',
                'Ã¢â‚¬Â³'   :'”',
                'Ã¢Ë†Â«'    :'∫', 
                'Ã¢Ë†Ëœ'    :'°',  
                'Ã¢Ë†Ë†'    :'∈',            
                'Ã¢Ë†â€˜'   :'Σ',                
                'Ã¢â‚¬Â²'   :"'", 
                "Ã¢â'¬Â²"   :"'",
                "Ã¢â‚¬Ëœ"   :"'",
                "Ã¢â‚¬â€"   :'—',
                "Ã¢â€°Â²"   :'≲',             
                '''Ã¢â'¬"''':'–',
                "Ã¢†'"      :'→',
                "ÃƒÂ¬"      :'μ',
                "ÃƒÂ¡"      :'á',   
                "ÃƒÂ©"      :'é',
                'ÃŽÂ¨'      :'Ψ',
                'ÃŽÂ³'      :'γ',
                'ÃŽÂ²'      :'β',
                'ÃŽÂ´'      :'δ',
                'ÃŽÂ»'      :'λ',               
                'Ã¢‰Â¡'     :'≡',
                'Ã¢—Â¡'     :'⊙',
                'Ã¢Ë†Â¼'    :'∼',
                'Ã¢Ë†Å¾'    :'∞',
                'Ã¢‰Ë†'     :'≈',
                'ÃŽÂ¼'      :'μ',
                'â‹…'       :'.',                           
                'ÃŽâ€'      :'Δ',                
                'ÃŽÂµ'      :'ε',
                'Ã¢‰Â¤'     :'≤',
                'âŠ™'       :'⊙',
                'ÃŽÂ©'      :'Ω',
                'ÃŽÂ±'      :'α',
                'ÃŽ'        :'α',
                "âˆ'"       :'−',
                'â†¦'       :'↦',
                'âˆ‹'       :'∋',  
                'âˆˆ'       :'∈',   
                'âˆ—'       :'*',
                "Ã¢‰Æ'"     :'≃',
                'Ã¢Å â"¢'   :'⊙',
                'hÌ£'       :'ḥ',
                'iÌ„'       :'ī',
                'aÌ„'       :'ā',
                'Ã‡'        :'Ç',
                'Ã¼'        :'ü',
                'Ã¢â€ â€™'  :'→',
                'Â©'        :'©',
                'Ãâ€ž'     :'τ',  
                'Ã¡Â¸Å¸'    :'ḟ',
                'Ãâ€¡'     :'χ',
                'Ãâ€¢'     :'Φ', 
                'ÃˆÂ¯'      :'⊙', 
                'ÃÆ’'      :'σ',
                'Ïƒ'        :'σ',
                'ÃŠËœ'      :'⊙',
                'ÃƒÂ¢ÅÂ Ã¢"Â¢' : '⊙',
                'Ã¢Ë†â€™'   :'-',
                'Ãï¿½'     :'ρ',   
                'Î£'        :'Σ',       
                'á¸¢'       : 'Ḣ',
                'Á¸¢'       : 'Ḣ',
                'â€œ'       : '“',
                'â€'       : '”',
                'â€™'       : '’',
                'â€˜'       : '‘',
                'â€”'       : '–',
                'â€“'       : '—',
                'â€¢'       : '-',
                'â€¦'       : '…',
                'â‰¥'       : '≥',
                'â‰¤'       : '≤',
                'â†’'       : '→',
                'âˆž'       : '∞', 
                'âˆ‚'       : '∂',
                'â‹…'       :'∙',
                'âˆ†'       :'Δ',
                'Ã‚Â»'      :'"',
                'Ã‚Â'       :' ',
                'Æ’'        :'ƒ',
                '/&'        :'&',
                '/%'        :'%',
                'Â°'        :'°',
                'Ã—'        :'x',
                'Ãª'        :'ê',
                'ÅŸ'        :'ş',
                'Î¼'        :'µ',
                '×'         :'x',
                'ÃŸ'        :'ß',
                'Ã©'        :'é',
                'Ã¶'        :'ö',
                'ã¶'        :'ö',
               	'â‰¡'       :'≡',
                'Â±'        :'±',                
                'Ã…'        :'Å',
                'Ã¥'        :'Å',
                'pÌ"'       :'p̅',
                'Î»'        :'λ',
                'Î³'        :'γ',                
                'Î±'        :'α',
                'Î”'        :'Δ',  
                'Ï‰'        :'ω',
                'Ï€'        :'π',
                'Ï„'        :'τ',
                'Ïˆ'        :'ψ', 
                'Ïµ'        :'ε',
                'Îµ'        :'∈',
                'Ã­'         :'í',
                '˜'         :'~',
                'Ã¤'        :'ä',
                'Â€œ'       :'"',
                'â€'       :'"',
                'Ã‰'        :'É',
                'Ãƒ–'       :'×',
                'Ã‚'        :'',
                'Ã¯Â»Â¿'    :'',
                'Ó§'        :'ö'                                
                }


def special_char_remove( col ):
    global special_char
    # I think the following characters are produced from the text editor converted our UTF-8 
    #   characters into some other character set, like ISO-8859-1.
    for i in range( len(col) ):
        col[i] = str(col[i])
        for char in special_char:
            while char in col[i]:
                col[i] = col[i].replace(char, special_char[char])
        col[i] = ftfy.fix_text(col[i])

    # For some reaseon ftfy.fix_text(.) messes with some of the characters we replaced
    #  however, the first time (above) is needed for the ftfy to work.
    for i in range( len(col) ):
        for char in special_char:
            while char in col[i]:
                col[i] = col[i].replace(char, special_char[char])

    # for i in ['Ã','¥','¶','â','Â','¼','½','¾']:
    #     if i in string:
    #         with open('zzz_REVIEW_STRING.txt','a+') as f:
    #             try:
    #                 f.write(f"REVIEW STRING in {excelName} {cell}:\n {string}\n\n")
    #             except:
    #                 print(f"REVIEW STRING in {excelName} {cell}:\n {string}\n\n")
    #         break
    return col


def ensure_encryption(df):

    for col in ["Funding Sponsor", "keywords", "abstract", "publisher", "source_publication"]:
        # try:
        df.loc[ :, col ] = df.loc[ :, col ].fillna('').values
        df.loc[ :, col ] = special_char_remove( df[ col ].values )
        # except TypeError as e:
        #     print( e.__str__() )
        #     print( df.loc[ :, col ] )
    
    for col in df.columns:
        if col[-6:] in ["_lname", "_mname", "_fname"]:
            df.loc[ :, col ] = df.loc[ :, col ].fillna('').values
            df.loc[ :, col ] = special_char_remove( df[ col ].values )

    return df