function beer{
    param($num)
    

    For ($num; $num -gt 0; $num --) {

        $new = ($num-1)
     
        $string = " bottles of beer on the wall!"

            If ($num -gt 1)
            {
            $result = "$num bottles of beer on the wall, $num bottles of beer, Take one down, pass it around, " + ($new) +, $string 
            }

            Elseif ($num -lt 2 )
            {
             $result = "$num bottle of beer on the wall, $num bottle of beer, Take one down, pass it around, " + ($new) +, " bottles of beer on the wall" 
            }
  
        Write-Host $result
    }
}
beer 3
