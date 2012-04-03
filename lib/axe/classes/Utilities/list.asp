<%

' File: list.asp
' 
' AXE(ASP Xtreme Evolution) list utility.
' 
' License:
' 
' This file is part of ASP Xtreme Evolution.
' Copyright (C) 2007-2012 Fabio Zendhi Nagao
' 
' ASP Xtreme Evolution is free software: you can redistribute it and/or modify
' it under the terms of the GNU Lesser General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
' 
' ASP Xtreme Evolution is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU Lesser General Public License for more details.
' 
' You should have received a copy of the GNU Lesser General Public License
' along with ASP Xtreme Evolution. If not, see <http://www.gnu.org/licenses/>.



' Class: List
' 
' This is an enhanced implementation of a doubly-linked list. It features
' additional methods which turns to be very useful to make queues, stacks, 
' dynamic allocation arrays and others.
' 
' In a doubly-linked list, each node contains, besides the next-node link, a
' second link field pointing to the previous node in the sequence. The two links
' may be called forward(s) and backwards.
' 
' About:
' 
'     - This implementation uses sentinels.
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ August 2009
' 
class List
    
    ' Property: classType
    ' 
    ' Class type.
    ' 
    ' Contains:
    ' 
    '   (string) - type
    ' 
    public classType

    ' Property: classVersion
    ' 
    ' Class version.
    ' 
    ' Contains:
    ' 
    '   (float) - version
    ' 
    public classVersion
    
    ' Property: count
    ' 
    ' Nodes count.
    ' 
    ' Contains:
    ' 
    '     (int) - Nodes count
    ' 
    public count
    
    ' Property: Head
    ' 
    ' List head reference.
    ' 
    ' Contains:
    ' 
    '     (List_Node) - List head reference
    ' 
    public Head
    
    ' Property: Foot
    ' 
    ' List foot referece.
    ' 
    ' Contains:
    ' 
    '     (List_Node) - List foot reference
    ' 
    public Foot
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.1.0.0"
        
        count = 0
        
        set Head = new List_Node
        set Foot = new List_Node
        
        set Head.pNext = Foot
        set Foot.pPrev = Head
    end sub
    
    private sub Class_terminate()
        dim Node : set Node = Head
        while(Node.hasNext)
            set Node = Node.pNext
            if( isObject(Node.data) ) then
                set Node.data = nothing
            end if
            set Node.pPrev = nothing
        wend
        set Node = nothing
        
        set Foot = nothing
        set Head = nothing
    end sub
    
    ' Function: items
    ' 
    ' Builds an enumerator to the list nodes.
    ' 
    ' Returns:
    ' 
    '     (mixed[]) - enumerator
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim L, greet : set L = new List
    ' 
    ' L.push("Welcome")
    ' L.push("歡迎光臨")
    ' L.push("Bienvenido")
    ' L.push("Bem vindo")
    ' L.push("환영합니다")
    ' L.push("ようこそ")
    ' 
    ' for each greet in L.items()
    '     Response.write( greet & vbNewline )
    ' next
    ' 
    ' set L = nothing
    ' 
    ' (end code)
    ' 
    public function items()
        dim a, i, Node
        
        redim a(count - 1)
        i = 0
        
        set Node = Head
        while(Node.hasNext)
            set Node = Node.pNext
            
            if( isObject(Node.data) ) then
                set a(i) = Node.data
            else
                a(i) = Node.data
            end if
            
            i = i + 1
        wend
        set Node = nothing
        
        items = a
    end function
    
    ' Function: unshift
    ' 
    ' Adds one element to the beginning of the list and returns the new length.
    ' 
    ' Parameters:
    ' 
    '     (mixed) - Element to be inserted
    ' 
    ' Returns:
    ' 
    '     (int) - The current number of elements
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim L : set L = new List
    ' 
    ' L.unshift("Welcome")
    ' L.unshift("歡迎光臨")
    ' L.unshift("Bienvenido")
    ' L.unshift("Bem vindo")
    ' L.unshift("환영합니다")
    ' L.unshift("ようこそ")
    ' 
    ' Response.write(L.pop() & vbNewline) ' prints "Welcome"
    ' Response.write(L.count & vbNewline) ' prints 5
    ' 
    ' set L = nothing
    ' 
    ' (end code)
    ' 
    public function unshift(mixed)
        dim Node : set Node = new List_Node
        
        if( isObject(mixed) ) then
            set Node.data = mixed
        else
            Node.data = mixed
        end if
        
        set Node.pNext = Head.pNext
        set Node.pPrev = Head
        
        set Head.pNext.pPrev = Node
        set Head.pNext = Node
        
        set Node = nothing
        
        count = count + 1
        unshift = count
    end function
    
    ' Function: shift
    ' 
    ' Removes and returns the first element of the list.
    ' 
    ' Returns:
    ' 
    '     (mixed) - Removed element
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim L : set L = new List
    ' 
    ' L.unshift("Welcome")
    ' L.unshift("歡迎光臨")
    ' L.unshift("Bienvenido")
    ' L.unshift("Bem vindo")
    ' L.unshift("환영합니다")
    ' L.unshift("ようこそ")
    ' 
    ' Response.write(L.shift() & vbNewline) ' prints "ようこそ" (yōkoso)
    ' Response.write(L.count & vbNewline) ' prints 5
    ' 
    ' set L = nothing
    ' 
    ' (end code)
    ' 
    public function shift()
        if(Head.hasNext) then
            set shift = Head.pNext
            count = count - 1
            
            set Head.pNext = shift.pNext
            set shift.pNext.pPrev = Head
            
            shift.pPrev = empty
            shift.pNext = empty
            
            if( isObject( shift.data ) ) then
                set shift = shift.data
            else
                shift = shift.data
            end if
        else
            shift = empty
        end if
    end function
    
    ' Function: push
    ' 
    ' Adds one element to the end of the list and returns the new length.
    ' 
    ' Parameters:
    ' 
    '     (mixed) - Element to be inserted
    ' 
    ' Returns:
    ' 
    '     (int) - The current number of elements
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim L : set L = new List
    ' 
    ' L.push("Welcome")
    ' L.push("歡迎光臨")
    ' L.push("Bienvenido")
    ' L.push("Bem vindo")
    ' L.push("환영합니다")
    ' L.push("ようこそ")
    ' 
    ' Response.write(L.shift() & vbNewline) ' prints "Welcome"
    ' Response.write(L.count & vbNewline) ' prints 5
    ' 
    ' set L = nothing
    ' 
    ' (end code)
    ' 
    public function push(mixed)
        dim Node : set Node = new List_Node
        
        if( isObject(mixed) ) then
            set Node.data = mixed
        else
            Node.data = mixed
        end if
        
        set Node.pNext = Foot
        set Node.pPrev = Foot.pPrev
        
        set Foot.pPrev.pNext = Node
        set Foot.pPrev = Node
        
        set Node = nothing
        
        count = count + 1
        push = count
    end function
    
    ' Function: pop
    ' 
    ' Removes and returns the last element of the list.
    ' 
    ' Returns:
    ' 
    '     (int) - The current number of elements
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim L : set L = new List
    ' 
    ' L.push("Welcome")
    ' L.push("歡迎光臨")
    ' L.push("Bienvenido")
    ' L.push("Bem vindo")
    ' L.push("환영합니다")
    ' L.push("ようこそ")
    ' 
    ' Response.write(L.pop() & vbNewline) ' prints "ようこそ" (yōkoso)
    ' Response.write(L.count & vbNewline) ' prints 5
    ' 
    ' set L = nothing
    ' 
    ' (end code)
    ' 
    public function pop()
        if(Foot.hasPrev) then
            set pop = Foot.pPrev
            count = count - 1
            
            set Foot.pPrev = pop.pPrev
            set pop.pPrev.pNext = Foot
            
            pop.pPrev = empty
            pop.pNext = empty
            
            if( isObject( pop.data ) ) then
                set pop = pop.data
            else
                pop = pop.data
            end if
        else
            pop = empty
        end if
    end function
    
    ' Function: search
    ' 
    ' Returns an element based on the result of the assert.
    ' 
    ' Parameters:
    ' 
    '     (assert|string) - Assert function or Assert function name
    ' 
    ' Returns:
    ' 
    '     (List_Node) - Matching Node
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim L : set L = new List
    ' 
    ' L.push("Welcome")
    ' L.push("歡迎光臨")
    ' L.push("Bienvenido")
    ' L.push("Bem vindo")
    ' L.push("환영합니다")
    ' L.push("ようこそ")
    ' 
    ' function bienvenidoDetector(Node)
    '     bienvenidoDetector = false
    '     if(Node.data = "Bienvenido") then bienvenidoDetector = true
    ' end function
    ' 
    ' Response.write( L.search("bienvenidoDetector") )' prints "Bienvenido"
    ' 
    ' set L = nothing
    ' 
    ' (end code)
    ' 
    public function search(assert)
        if(lcase(typename(assert)) <> "object") then set assert = getRef(assert)
        
        dim Node : set Node = Head
        while(Node.hasNext)
            set Node = Node.pNext
            if(assert(Node)) then
                set search = Node
                set Node = nothing
                exit function
            end if
        wend
        set Node = nothing
    end function
    
    ' Function: remove
    ' 
    ' Removes an element based on the result of the assert.
    ' 
    ' Parameters:
    ' 
    '     (assert|string) - Assert function or Assert function name
    ' 
    ' Returns:
    ' 
    '     (int) - The current number of elements
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim L : set L = new List
    ' 
    ' L.push("Welcome")
    ' L.push("歡迎光臨")
    ' L.push("Bienvenido")
    ' L.push("Bem vindo")
    ' L.push("환영합니다")
    ' L.push("ようこそ")
    ' 
    ' function bienvenidoDetector(Node)
    '     bienvenidoDetector = false
    '     if(Node.data = "Hello World") then bienvenidoDetector = true
    ' end function
    ' 
    ' Response.write( L.remove("bienvenidoDetector") )' prints 5
    ' 
    ' set L = nothing
    ' 
    ' (end code)
    ' 
    public function remove(assert)
        dim Node : set Node = search(assert)
        set Node.pPrev.pNext = Node.pNext
        set Node.pNext.pPrev = Node.pPrev
        set Node = nothing
        count = count - 1
        remove = count
    end function
    
end class



' Class: List_Node
' 
' Node implementation for a Doubly Linked List with sentinels.
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ August 2009
' 
class List_Node
    
    ' Property: classType
    ' 
    ' Class type.
    ' 
    ' Contains:
    ' 
    '   (string) - type
    ' 
    public classType

    ' Property: classVersion
    ' 
    ' Class version.
    ' 
    ' Contains:
    ' 
    '   (float) - version
    ' 
    public classVersion
    
    ' Property: data
    ' 
    ' Node content
    ' 
    ' Contains:
    ' 
    '     (mixed) - Node data
    ' 
    public data
    
    ' Property: pPrev
    ' 
    ' Previous node reference.
    ' 
    ' Contains:
    ' 
    '     (List_Node) - Previous node reference
    ' 
    public pPrev
    
    ' Property: pNext
    ' 
    ' Next node reference.
    ' 
    ' Contains:
    ' 
    '     (List_Node) - Next node reference
    ' 
    public pNext
    
    ' Property: hasPrev
    ' 
    ' Informs if you can keep going backwards from the current node.
    ' 
    ' Contains:
    ' 
    '     (true)  - you can
    '     (false) - otherwise
    ' 
    public property get hasPrev
        hasPrev = true
        if(isEmpty(Me.pPrev)) then hasPrev = false : exit property
        if(isEmpty(Me.pPrev.pPrev)) then hasPrev = false
    end property
    
    ' Property: hasNext
    ' 
    ' Informs if you can keep going forward from the current node.
    ' 
    ' Contains:
    ' 
    '     (true)  - you can
    '     (false) - otherwise
    ' 
    public property get hasNext
        hasNext = true
        if(isEmpty(Me.pNext)) then hasNext = false : exit property
        if(isEmpty(Me.pNext.pNext)) then hasNext = false
    end property
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0.0"
    end sub
    
    private sub Class_terminate()
    end sub
    
end class

%>
