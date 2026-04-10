option compare database
option explicit

function pi() as double
on error goto err_handler

pi = 4 * atn(1)

exit function
err_handler:
    call handleerror("modMath", "pi", err.description, err.number)
end function

function asin(x) as double
on error goto err_handler

select case x
    case 1
        asin = pi / 2
    case -1
        asin = (3 * pi) / 2
    case else
        asin = atn(x / sqr(-x * x + 1))
end select

exit function
err_handler:
    call handleerror("modMath", "Asin", err.description, err.number)
end function

function acos(x) as double
on error goto err_handler

select case x
    case 1
        acos = 0
    case -1
        acos = pi
    case else
        acos = atn(-x / sqr(-x * x + 1)) + 2 * atn(1)
end select

exit function
err_handler:
    call handleerror("modMath", "Acos", err.description, err.number)
end function

function gramstolbs(gramsvalue) as double
on error goto err_handler

gramstolbs = gramsvalue * 0.00220462

exit function
err_handler:
    call handleerror("modMath", "gramsToLbs", err.description, err.number)
end function

function randomnumber(low as long, high as long) as long
on error goto err_handler

randomize
randomnumber = int((high - low + 1) * rnd() + low)

exit function
err_handler:
    call handleerror("modMath", "randomNumber", err.description, err.number)
end function
