
[bits 32]    
    push edi                    ;Preserve edi

    mov edi, [esp+12]           ;Copy dest pointer into destination index
    mov eax, [esp+16]           ;Byte count in eax
    mov edx, [esp+20]           ;Copy fill character

; fill all four bytes of ebx with the value by shifting and or-ing
    mov 	ecx, edx
    shl 	ecx, 8
    or 	edx, ecx
    mov     ecx, edx
    shl 	edx, 16
    or 	edx, ecx

; align the destination on an 16 byte boundary
    mov     ecx, edi
    and     ecx, 15
    sub     eax, ecx
    
    jecxz 	_aligned
_align:    
    mov 	BYTE [edi],	dl
    add 	edi, 1
    sub     ecx, 1
    jnz    _align
    
_aligned:
    mov     ecx, eax
    and     eax, 15
    shr     ecx, 4
    jz	_tail1

; fill all 16 bytes of xmm0 with the fill value by pushing ebx to the stack four times and 
; then moving from memory to to the SSE register... is there a better way?
    emms
    push    edx
    push    edx
    push    edx
    push    edx
    movdqa  xmm0, [esp]
    add     esp, 16
_loop_sse:
    movq	[edi], mm0
    add     edi, 16
    sub     ecx, 1
    jnz    _loop_sse
    emms
    
_tail1:
    mov     ecx, eax
    jecxz 	_end
_tail2:    
    mov 	BYTE [edi],	dl
    add 	edi, 1
    sub     ecx, 1
    jnz    _tail2

_end:
    pop     edi
    ret     16