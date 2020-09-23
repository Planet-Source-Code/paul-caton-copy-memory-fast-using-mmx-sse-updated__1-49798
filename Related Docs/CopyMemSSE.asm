[bits 32]    
    push esi                    ;Preserve esi
    push edi                    ;Preserve edi

    mov edi, [esp+16]           ;Copy dest pointer into destination index
    mov esi, [esp+20]           ;Copy source pointer into source index
    mov ecx, [esp+24]           ;Byte count in ecx
    
    mov     eax, ecx
    sub     ecx, edi
    sub     ecx, eax
    and     ecx, 15
    sub     eax, ecx
    jle     l2

    emms
    rep     movsb
    mov     ecx, eax
    and     eax, 15
    shr     ecx, 4
    jz	l2
    sub	edi, esi
l1: movdqa	xmm0, [esi]
    movdqa	[edi+esi], xmm0
    add     esi, 16
    dec     ecx
    jnz     l1
    emms
    add	edi, esi
    
l2: add     ecx, eax
    rep     movsb

    pop     edi
    pop     esi
    ret     16