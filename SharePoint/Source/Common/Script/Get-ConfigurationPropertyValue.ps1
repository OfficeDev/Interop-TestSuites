#-------------------------------------------------------------------------
# Copyright (c) 2014 Microsoft Corporation. All rights reserved.
# Use of this sample source code is subject to the terms of the Microsoft license 
# agreement under which you licensed this sample source code and is provided AS-IS.
# If you did not accept the terms of the license agreement, you are not authorized 
# to use this sample source code. For the terms of the license, please see the 
# license agreement between you and Microsoft.
#-------------------------------------------------------------------------

param(
[string]$propertyName
)

$propertyValueVariable = gv "ptfprop$propertyName" -ea SilentlyContinue
if ($propertyValueVariable -ne $null)
{
    $regex = [regex] "\[(?<property>[^\[]+?)\]"
    if ($regex.IsMatch($propertyValueVariable.Value))
    {
        $matchEvaluator = [System.Text.RegularExpressions.MatchEvaluator]{
            $matchedPropertyName = $args[0].Groups["property"].Value
            $matchedPropertyValueVariable = gv "ptfprop$matchedPropertyName" -ea SilentlyContinue
            if($matchedPropertyValueVariable -ne $null)
            {
                return $matchedPropertyValueVariable.Value
            }
            else
            {
                return $args[0].Value
            }
        }
        
        return $regex.Replace($propertyValueVariable.Value, $matchEvaluator)
    }

    return $propertyValueVariable.Value
}
elseif ($propertyName -ieq "CommonConfigurationFileName")
{
    return $null
}
else
{
    throw "Property '$propertyName' was not found in the ptfconfig file. Note: When processing property values, string in square brackets ([...]) will be replaced with the property value whose name is the same string."
}