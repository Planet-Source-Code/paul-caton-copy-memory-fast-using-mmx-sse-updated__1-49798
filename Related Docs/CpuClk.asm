; Write the CPU clock cycle count (since cpu reset) into the passed ByRef Currency (64 bit) parameter

[bits 32]
    rdtsc                                   ;# Read the cpu clock cycle count into eax/edx
    mov ecx, dword [esp+8]                  ;# Address of Currency parameter into ecx
    mov [ecx], eax                          ;# Put eax to the low part of the currency variable
    mov [ecx+4], edx                        ;# Put edx to the high part of the currency variable
    xor eax, eax                            ;# Clear eax, tell VB a-okay
    ret 8                                   ;# Return
