;Determine the processor's SIMD support
; return: 0 = none
;         1 = MMX
;         2 = SSE
;         3 = SSE2
;         4 = SSE3
[BITS 32]

;Check for CPUID support
    pushfd
    pop	eax
    btc	eax, 15h
    push	eax
    popfd
    pushfd
    pop	edx
    xor	eax, edx
    jz	_cpuid_supported
    xor     eax, eax
    jmp     _set_results2

_cpuid_supported:
    push    ebx             ;Preserve ebx
    mov     eax, 1h         ;CPUID level 1
    cpuid                   ;CPUID
    xor     eax, eax
    bt      edx, 23         ;MMX
    jnc     _set_results1
    inc     eax
    bt      edx, 25         ;SSE
    jnc     _set_results1
    inc     eax
    bt      edx, 26         ;SSE2
    jnc     _set_results1
    inc     eax
    bt      ecx, 0          ;SSE3
    jnc     _set_results1
    inc     eax
_set_results1:
    pop     ebx             ;Restore ebx
_set_results2
    mov     edx, [esp+08h]  ;Get the address of the return value
    mov     [edx], eax      ;Write the result
    xor     eax, eax        ;Clear eax - tells VB that all is okay
    ret     8               ;Return