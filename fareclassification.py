def returnFare(wkUsageClassificationCd, wkPrice):
    if wkUsageClassificationCd == "i09-001s":
        yield 3
        yield "i09-001s"
    elif wkUsageClassificationCd == "i09-002s":
        if wkPrice < 10000:
            yield 3
            yield "i09-002s"
        else:
            yield 2
            yield "i09-100a"
    elif wkUsageClassificationCd == "i09-003s":
        if wkPrice < 10000:
            yield 3
            yield "i09-003s"
        else:
            yield 2
            yield "i09-100a"
    elif wkUsageClassificationCd == "i09-004s":
            yield 3
            yield "i09-004s"
    elif wkUsageClassificationCd == "i09-005s":
            yield 3
            yield "i09-005s"
    elif wkUsageClassificationCd == "i09-006s":
            yield 2
            yield "i09-006s"
    elif wkUsageClassificationCd == "i09-007s":
            yield 2
            yield "i09-007s"
    elif wkUsageClassificationCd == "i09-008a":
            yield 3
            yield "i09-008a"
    elif wkUsageClassificationCd == "i09-009a":
        if wkPrice < 10000:
            yield 3
            yield "i09-009a"
        else:
            yield 2
            yield "i09-100b"
    elif wkUsageClassificationCd == "i09-010a":
        if wkPrice < 10000:
            yield 3
            yield "i09-010a"
        else:
            yield 2
            yield "i09-100b"
    elif wkUsageClassificationCd == "i09-011a":
            yield 3
            yield "i09-011a"
    elif wkUsageClassificationCd == "i09-012a":
            yield 3
            yield "i09-012a"
    elif wkUsageClassificationCd == "i09-013a":
            yield 2
            yield "i09-013a"
    elif wkUsageClassificationCd == "i09-014a":
            yield 2
            yield "i09-014a"
    elif wkUsageClassificationCd == "i09-001d":
            yield 3
            yield "i09-001d"
    elif wkUsageClassificationCd == "i09-002d":
        if wkPrice < 10000:
            yield 3
            yield "i09-002d"
        else:
            yield 2
            yield "i09-100a"
    elif wkUsageClassificationCd == "i09-003d":
        if wkPrice < 10000:
            yield 3
            yield "i09-003d"
        else:
            yield 2
            yield "i09-100a"
    elif wkUsageClassificationCd == "i09-004d":
            yield 2
            yield "i09-004d"
    elif wkUsageClassificationCd == "i09-005d":
            yield 3
            yield "i09-005d"
    elif wkUsageClassificationCd == "i09-006d":
            yield 2
            yield "i09-006d"
    elif wkUsageClassificationCd == "i09-007d":
            yield 2
            yield "i09-007d"
    else:
            yield 3
            yield "i09-002d"
