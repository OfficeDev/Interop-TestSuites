#include "nspi.h"
#include "wchar.h"

void __RPC_FAR * __RPC_USER midl_user_allocate(size_t cBytes)
{
    return((void __RPC_FAR *) malloc(cBytes));
}

void __RPC_API midl_user_free(void __RPC_FAR * p)
{
    free(p);
}