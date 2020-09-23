[bits 32]    
    push esi                    ;Preserve esi
    push edi                    ;Preserve edi

    mov edi, [esp+16]           ;Copy dest pointer into destination index
    mov esi, [esp+20]           ;Copy source pointer into source index
    mov ecx, [esp+24]           ;Byte count in ecx
    
    mov     eax, ecx
    sub     ecx, edi
    sub     ecx, eax
    and     ecx, 7
    sub     eax, ecx
    jle     l2
    emms
    rep     movsb
    mov     ecx, eax
    and     eax, 7
    shr     ecx, 3
    jz	l2
    sub	edi, esi
l1: movq	mm0, [esi]
    movq	[edi+esi], mm0
    add     esi,8
    dec     ecx
    jnz     l1
    add	edi, esi
    emms
l2: add     ecx, eax
    rep     movsb

    pop     edi
    pop     esi
    ret     16