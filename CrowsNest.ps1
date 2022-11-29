

Function message {
    Param ($item)

    if($item.startsWith("a") -or $item.startsWith("e") -or $item.startsWith("i") -or $item.startsWith("o") -or $item.startsWith("u")){
    
    $value = "an"
    }
    else{$value="a"}

    Write-host ("Captain, " + $value + " " + $item + " off the lardboard bow")
    

}
Message "owl"