def returnFare(wkUsageClassificationCd):
    if wkUsageClassificationCd == "i08-001d":
        yield 2
        yield "i08-001d"
    elif wkUsageClassificationCd == "i08-002g":
        yield 2
        yield "i08-002d"
    elif wkUsageClassificationCd == "i08-003g":
        yield 2
        yield "i08-003d"
    elif wkUsageClassificationCd == "i08-004g":
        yield 2
        yield "i08-004d"
    elif wkUsageClassificationCd == "i08-001g":
        yield 2
        yield "i08-001g"
    elif wkUsageClassificationCd == "i08-002g":
        yield 2
        yield "i08-002g"
    elif wkUsageClassificationCd == "i08-003g":
        yield 2
        yield "i08-003g"
    elif wkUsageClassificationCd == "i08-004g":
        yield 2
        yield "i08-004g"
    else:
        yield 2
        yield "i08-004sk"
