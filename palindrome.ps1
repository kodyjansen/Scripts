Function palindrome
{
	Param ($word)
	$targetword =$word
	

	$word1 = $targetword.substring(0, $targetword.indexof(" "))
	
	$word2 = $targetword.substring($targetword.indexof(" ") + 1, ($targetword.length - $targetword.indexof(" ") - 1))
	

	$word2 + " " + $word1
	#$result = $result + $word.substring($num-1, 1)
		
	


	
}
palindrome "banana grams"
