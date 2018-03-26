<#\G.{2} means split the string 2 by 2 characters (19 is the total number of characters)
[char][int]"0x$_" --> means "Do a foreach loop to convert each couple of characters to hexadecimal"
\G specifies that the match occur at the point where the previous match ended. When used with Match.NextMatch(),
this ensures that matches are all contiguous
(?<= ) Zero-width positive look behind assertion. Continues match only if the subexpression matches at this position
on the left. For example, (?<=19)99 matches instances of 99 that follow 19.
#>

-join ("796F75722D656D61696C40646F6D61696E2E636F6D" -split"(?<=\G.{2})"|where{$_}|%{[char][int]"0x$_"})