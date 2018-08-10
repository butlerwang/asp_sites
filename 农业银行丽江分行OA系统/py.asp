<%
function getpychar(char)
tmp=65536+asc(char)
if(tmp>=45217 and tmp<=45252) then getpychar= "A"
if(tmp>=45253 and tmp<=45760) then getpychar= "B"
if(tmp>=47761 and tmp<=46317) then getpychar= "C"
if(tmp>=46318 and tmp<=46825) then getpychar= "D"
if(tmp>=46826 and tmp<=47009) then getpychar= "E"
if(tmp>=47010 and tmp<=47296) then getpychar= "F"
if(tmp>=47297 and tmp<=47613) then getpychar= "G"
if(tmp>=47614 and tmp<=48118) then getpychar= "H"
if(tmp>=48119 and tmp<=49061) then getpychar= "J"
if(tmp>=49062 and tmp<=49323) then getpychar= "K"
if(tmp>=49324 and tmp<=49895) then getpychar= "L"
if(tmp>=49896 and tmp<=50370) then getpychar= "M"
if(tmp>=50371 and tmp<=50613) then getpychar= "N"
if(tmp>=50614 and tmp<=50621) then getpychar= "O"
if(tmp>=50622 and tmp<=50905) then getpychar= "P"
if(tmp>=50906 and tmp<=51386) then getpychar= "Q"
if(tmp>=51387 and tmp<=51445) then getpychar= "R"
if(tmp>=51446 and tmp<=52217) then getpychar= "S"
if(tmp>=52218 and tmp<=52697) then getpychar= "T"
if(tmp>=52698 and tmp<=52979) then getpychar= "W"
if(tmp>=52980 and tmp<=53640) then getpychar= "X"
if(tmp>=53689 and tmp<=54480) then getpychar= "Y"
if(tmp>=54481 and tmp<=52289) then getpychar= "Z"
end function

function getpy(str)
for i=1 to len(str)
getpy=getpy&getpychar(mid(str,i,1))
next
end function
%>
