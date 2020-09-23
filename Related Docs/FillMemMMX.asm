
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

; align the destination on an 8 byte boundary
    mov     ecx, edi
    and     ecx, 7
    sub     eax, ecx
    
    jecxz 	_aligned
_align:    
    mov 	BYTE [edi],	dl
    add 	edi, 1
    sub     ecx, 1
    jnz    _align
    
_aligned:
    mov     ecx, eax
    and     eax, 7
    shr     ecx, 3
    jz	_tail1

; fill all eight bytes of mm0 with the fill value by pushing ebx to the stack twice and 
; then moving from memory to to the MMX register... is there a better way?
    emms
    push    edx
    push    edx
    movq    mm0, [esp]
    add     esp, 8
_loop_mmx:
    movq	[edi], mm0
    add     edi, 8
    sub     ecx, 1
    jnz    _loop_mmx
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