#include "StdAfx.h"

CheckInput::CheckInput(void)
{
}
//numeric
bool CheckInput::IsNumeric(System::Object^ oTarget)
{
    double dNullable;
    return System::Double::TryParse(
        oTarget->ToString(),
        System::Globalization::NumberStyles::Any,
        nullptr,
        dNullable
    );
}
//sizes
bool CheckInput::IsWHD(System::Object^ oTarget)
{
    double dNullable;
    return System::Double::TryParse(
        oTarget->ToString(),
		System::Globalization::NumberStyles::AllowDecimalPoint,
		nullptr,
		dNullable
    );
}