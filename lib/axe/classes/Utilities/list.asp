<%

' File: list.asp
' 
' AXE(ASP Xtreme Evolution) list utility.
' 
' License:
' 
' This file is part of ASP Xtreme Evolution.
' Copyright (C) 2007-2009 Fabio Zendhi Nagao
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
        classVersion = "1.0.0.0"
        
        count = 0
        
        set Head = new List_Node
        set Foot = new List_Node
        
        set Head.pNext = Foot
        set Foot.pPrev = Head
    end sub
    
    private sub Class_terminate()
        set Foot = nothing
        set Head = nothing
    end sub
    
    ' Function: unshift
    ' 
    ' Adds one element to the beginning of the list and returns the new length.
    ' 
    ' Parameters:
    ' 
    '     (List_Node) - Element to be inserted
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
    ' dim Node : set Node = new List_node
    ' Node.data = "Hello World"
    ' Response.write(L.unshift(Node)) ' prints 1
    ' set Node = nothing
    ' 
    ' Response.write(L.Head.pNext.data) ' prints "Hello World"
    ' Response.write(L.Head.pNext.hasNext) ' prints false
    ' 
    ' set L = nothing
    ' 
    ' (end code)
    ' 
    public function unshift(Node)
        set Node.pNext = Head.pNext
        set Node.pPrev = Head
        
        set Head.pNext.pPrev = Node
        set Head.pNext = Node
        
        count = count + 1
        unshift = count
    end function
    
    ' Function: shift
    ' 
    ' Removes and returns the first element of the list.
    ' 
    ' Returns:
    ' 
    '     (List_Node) - Removed element
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim L : set L = new List
    ' 
    ' dim Node : set Node = new List_node
    ' Node.data = "Hello World"
    ' L.unshift(Node)
    ' Response.write(L.shift().data) ' prints "Hello World"
    ' Response.write(L.count) ' prints 0
    ' set Node = nothing
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
    '     (List_Node) - Element to be inserted
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
    ' dim Node : set Node = new List_node
    ' Node.data = "Hello World"
    ' Response.write(L.push(Node)) ' prints 1
    ' set Node = nothing
    ' 
    ' Response.write(L.Head.pNext.data) ' prints "Hello World"
    ' Response.write(L.Head.pNext.hasNext) ' prints false
    ' 
    ' set L = nothing
    ' 
    ' (end code)
    ' 
    public function push(Node)
        set Node.pNext = Foot
        set Node.pPrev = Foot.pPrev
        
        set Foot.pPrev.pNext = Node
        set Foot.pPrev = Node
        
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
    ' dim Node : set Node = new List_node
    ' Node.data = "Hello World"
    ' L.unshift(Node)
    ' Response.write(L.pop().data) ' prints "Hello World"
    ' Response.write(L.count) ' prints 0
    ' set Node = nothing
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
    ' dim Node : set Node = new List_node
    ' Node.data = "Hello World"
    ' Response.write(L.push(Node)) ' prints 1
    ' set Node = nothing
    ' 
    ' function helloWorldDetector(Node)
    '     helloWorldDetector = false
    '     if(Node.data = "Hello World") then helloWorldDetector = true
    ' end function
    ' 
    ' Response.write( L.search("helloWorldDetector").data )
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
    ' dim Node : set Node = new List_node
    ' Node.data = "Hello World"
    ' Response.write(L.push(Node)) ' prints 1
    ' set Node = nothing
    ' 
    ' function helloWorldDetector(Node)
    '     helloWorldDetector = false
    '     if(Node.data = "Hello World") then helloWorldDetector = true
    ' end function
    ' 
    ' Response.write( L.remove("helloWorldDetector") )
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
    '     (variant) - Node data
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
