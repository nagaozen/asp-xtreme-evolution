<%

' File: acl.asp
' 
' AXE(ASP Xtreme Evolution) implementation of RBAC utility.
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



' Class: Acl
' 
' Acl provides a flexible Role Based Access Control (RBAC) implementation for 
' privileges management. In general, an application may utilize this utility to 
' control access to certain protected objects by other requesting objects.
' 
' The features provided by this class are widely discussed and studied by the 
' National Institute of Standard and Technology (NIST). For the papers and more 
' info visit <http://csrc.nist.gov/groups/SNS/rbac/>.
' 
' Dependencies:
' 
'     - JSON2 class (/lib/axe/classes/Parsers/json2.asp)
' 
' About:
' 
'     - Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br> @ Dec 2010
' 
class Acl
    
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
    
    ' Property: [_Users]
    ' 
    ' {private} User -> Roles mappings.
    ' 
    ' Contains:
    ' 
    '     (Object) - in memory user -> roles mappings
    ' 
    private [_Users]
    
    ' Property: [_Roles]
    ' 
    ' {private} Roles hierarchy.
    ' 
    ' Contains:
    ' 
    '     (Object) - in memory roles hierarchy tree
    ' 
    private [_Roles]
    
    ' Property: [_Resources]
    ' 
    ' {private} Resources hierarchy.
    ' 
    ' Contains:
    ' 
    '     (Object) - in memory resources hierarchy tree
    ' 
    private [_Resources]
    
    ' Property: [_Rules]
    ' 
    ' {private} Rules between roles and resources.
    ' 
    ' Contains:
    ' 
    '     (Object) - in memory permission assignments between roles and resources
    ' 
    private [_Rules]
    
    ' Property: Media
    ' 
    ' Acl_Interface implementation
    ' 
    ' Contains:
    ' 
    '     (Acl_Interface) - Media implementing Acl_Interface
    ' 
    public Media
    
    ' Subroutine: [_ε]
    ' 
    ' {private} Checks for an media assignment.
    ' 
    private sub [_ε]
        if( isEmpty(Media) ) then
            Err.raise 5, "Evolved AXE runtime error", "Invalid procedure call or argument. Missing a Acl_Interface media."
        end if
    end sub
    
    private sub Class_initialize()
        classType    = typename(Me)
        classVersion = "1.0.0"
        
        set [_Users] = JSON.parse("{}")
        set [_Roles] = JSON.parse("{}")
        set [_Resources] = JSON.parse("{}")
        set [_Rules] = JSON.parse("{}")
    end sub
    
    private sub Class_terminate()
        set [_Rules] = nothing
        set [_Resources] = nothing
        set [_Roles] = nothing
        set [_Users] = nothing
    end sub
    
    ' Subroutine: addRole
    ' 
    ' Adds a role having an identifier unique to the roles registry.
    ' 
    ' Parameters:
    ' 
    '     (string) - role identifier
    '     (string) - role identifier or null
    ' 
    public sub addRole(id, parent)
        if( not isEmpty( [_Roles].get(id) ) ) then
            Err.raise 17, "Evolved AXE ACL runtime error", strsubstitute( _
                "Can't perform requested operation. The role id '{0}' already exists.", _
                array(id) _
            )
        end if
        
        if( not isNull(parent) ) then
            if( isEmpty( [_Roles].get(parent) ) ) then
                Err.raise 17, "Evolved AXE ACL runtime error", strsubstitute( _
                    "Can't perform requested operation. The role parent id '{0}' does not exists.", _
                    array(parent) _
                )
            end if
            
            call ( [_Roles].get(parent).get(1) ).set(id, null)
        end if
        
        call [_Roles].set( id, array( parent, JSON.parse("{}") ) )
    end sub
    
    ' Subroutine: remRole
    ' 
    ' Removes the role from the registry.
    ' 
    ' Parameters:
    ' 
    '     (string) - role identifier
    ' 
    public sub remRole(id)
        dim parent : parent = [_Roles].get(id).get(0)
        if( not isNull( parent ) ) then
            [_Roles].get(parent).get(1).delete(id)
        end if
        
        for each child in [_Roles].get(id).get(1).keys()
            call remRole( child )
        next
        
        call [_Roles].delete( id )
    end sub
    
    ' Subroutine: addResource
    ' 
    ' Adds a resource having an identifier unique to the resources registry.
    ' 
    ' Parameters:
    ' 
    '     (string) - resource identifier
    '     (string) - resource identifier or null
    ' 
    public sub addResource(id, parent)
        if( not isEmpty( [_Resources].get(id) ) ) then
            Err.raise 17, "Evolved AXE ACL runtime error", strsubstitute( _
                "Can't perform requested operation. The resource id '{0}' already exists.", _
                array(id) _
            )
        end if
        
        if( not isNull(parent) ) then
            if( isEmpty( [_Resources].get(parent) ) ) then
                Err.raise 17, "Evolved AXE ACL runtime error", strsubstitute( _
                    "Can't perform requested operation. The resource parent id '{0}' does not exists.", _
                    array(parent) _
                )
            end if
            
            call ( [_Resources].get(parent).get(1) ).set(id, null)
        end if
        
        call [_Resources].set( id, array( parent, JSON.parse("{}") ) )
    end sub
    
    ' Subroutine: remResource
    ' 
    ' Removes the resource from the registry.
    ' 
    ' Parameters:
    ' 
    '     (string) - resource identifier
    ' 
    public sub remResource(id)
        dim parent : parent = [_Resources].get(id).get(0)
        if( not isNull( parent ) ) then
            [_Resources].get(parent).get(1).delete(id)
        end if
        
        for each child in [_Resources].get(id).get(1).keys()
            call remResource( child )
        next
        
        call [_Resources].delete( id )
    end sub
    
    ' Subroutine: assign
    ' 
    ' Assigns role(s) to the user. NOTE: In the case of more than one role, the 
    ' roles list precedence works as a queue. (eg. array("Role_A", "Role_B"): if
    ' Role_A have a feature deny and Role_B have a feature allow, deny prevails)
    ' 
    ' Parameters:
    ' 
    '     (string)   - user identifier
    '     (string[]) - the chain of roles
    ' 
    public sub assign(id, roles)
        if( not isArray(roles) ) then
            roles = array(roles)
        end if
        
        if( isEmpty( [_Users].get(id) ) ) then
            call [_Users].set(id, JSON.parse("{}"))
        end if
        
        dim role : for each role in roles
            if( isEmpty( [_Roles].get(role) ) ) then
                Err.raise 17, "Evolved AXE ACL runtime error", strsubstitute( _
                    "Can't perform requested operation. The role id '{0}' does not exists.", _
                    array(role) _
                )
            end if
            call [_Users].get(id).set(role, null)
        next
    end sub
    
    ' Subroutine: unassign
    ' 
    ' Unnasigns role or roles to the user.
    ' 
    ' Parameters:
    ' 
    '     (string)   - user identifier
    '     (string[]) - the chain of roles
    ' 
    public sub unassign(id, roles)
        if( not isEmpty( [_Users].get(id) ) ) then
            if( not isArray(roles) ) then
                roles = array(roles)
            end if
            
            dim role : for each role in roles
                call [_Users].get(id).delete(role)
            next
            
            if( ubound( [_Users].get(id).keys() ) = -1 ) then
                call [_Users].delete(id)
            end if
        end if
    end sub
    
    ' Function: is
    ' 
    ' Checks if a user belongs to a role.
    ' 
    ' Parameters:
    ' 
    '     (string) - user identifier
    '     (string) - role identifier
    ' 
    ' Returns:
    ' 
    '     (boolean) - true, if he belongs; false otherwise.
    ' 
    public function [is](user, role)
        [is] = false
        if( not isEmpty( [_Users].get(user) ) ) then
            dim entry : for each entry in [_Users].get(user).keys()
                if( entry = role ) then
                    [is] = true
                    exit function
                end if
            next
        end if
    end function
    
    ' Subroutine: [_ƒ]
    ' 
    ' {private} Adds or removes an "allow" or "deny" rule to the ACL.
    ' 
    ' Parameters:
    ' 
    '     (string)   - role identifier
    '     (string)   - resource identifier
    '     (string)   - privilege identifier
    '     (function) - assert
    '     (string)   - action identifier
    '     (string)   - access identifier
    ' 
    private sub [_ƒ](role, resource, privilege, assert, action, access)
        if( isEmpty( [_Roles].get(role) ) ) then
            Err.raise 17, "Evolved AXE ACL runtime error", strsubstitute( _
                "Can't perform requested operation. The role id '{0}' does not exists.", _
                array(role) _
            )
        end if
        
        if( isEmpty( [_Resources].get(resource) ) ) then
            Err.raise 17, "Evolved AXE ACL runtime error", strsubstitute( _
                "Can't perform requested operation. The resource id '{0}' does not exists.", _
                array(resource) _
            )
        end if
        
        if( isEmpty( [_Rules].get(role) ) ) then
            call [_Rules].set(role, JSON.parse("{}"))
        end if
        
        if( isEmpty( [_Rules].get(role).get(resource) ) ) then
            call [_Rules].get(role).set(resource, JSON.parse("{}"))
        end if
        
        select case action
            case "ADD":
                call [_Rules].get(role).get(resource).set(privilege, array(access, assert))
            
            case "REM":
                call [_Rules].get(role).get(resource).delete(privilege)
        end select
        
        if( ubound( [_Rules].get(role).get(resource).keys() ) = -1 ) then
            call [_Rules].get(role).delete(resource)
        end if
        
        if( ubound( [_Rules].get(role).keys() ) = -1 ) then
            call [_Rules].delete(role)
        end if
    end sub
    
    ' Subroutine: allow
    ' 
    ' Adds an "allow" rule to the ACL.
    ' 
    ' Parameters:
    ' 
    '     (string)   - role identifier
    '     (string)   - resource identifier
    '     (string)   - privilege identifier
    '     (function) - assert
    ' 
    public sub allow(role, resource, privilege, assert)
        call [_ƒ](role, resource, privilege, assert, "ADD", "ALLOW")
    end sub
    
    ' Subroutine: remAllow
    ' 
    ' Removes an "allow" rule from the ACL.
    ' 
    ' Parameters:
    ' 
    '     (string)   - role identifier
    '     (string)   - resource identifier
    '     (string)   - privilege identifier
    ' 
    public sub remAllow(role, resource, privilege)
        call [_ƒ](role, resource, privilege, assert, "REM", "ALLOW")
    end sub
    
    ' Subroutine: deny
    ' 
    ' Adds a "deny" rule to the ACL.
    ' 
    ' Parameters:
    ' 
    '     (string)   - role identifier
    '     (string)   - resource identifier
    '     (string)   - privilege identifier
    '     (function) - assert
    ' 
    public sub deny(role, resource, privilege, assert)
        call [_ƒ](role, resource, privilege, assert, "ADD", "DENY")
    end sub
    
    ' Subroutine: remDeny
    ' 
    ' Removes a "deny" rule from the ACL.
    ' 
    ' Parameters:
    ' 
    '     (string) - role identifier
    '     (string) - resource identifier
    '     (string) - privilege identifier
    ' 
    public sub remDeny(role, resource, privilege)
        call [_ƒ](role, resource, privilege, assert, "REM", "DENY")
    end sub
    
    ' Function: [_φ]
    ' 
    ' {private} Evaluates an access against a roles x resources matrix.
    ' 
    ' Parameters:
    ' 
    '     (string[]) - the chain of roles
    '     (string[]) - the chain of resources
    '     (string)   - the privilege
    ' 
    ' Returns:
    ' 
    '     (mixed) - true, if an "allow" is found; false, if a "deny" is found; null otherwise
    ' 
    private function [_φ](roles, resources, privilege)
        [_φ] = null
        
        dim role, resource, assert
        
        for each role in roles
            for each resource in resources
                if( not isEmpty( [_Rules].get(role) ) ) then
                    if( not isEmpty( [_Rules].get(role).get(resource) ) ) then
                        if( not isEmpty( [_Rules].get(role).get(resource).get(privilege) ) ) then
                            if( isNull( [_Rules].get(role).get(resource).get(privilege).get(1) ) ) then
                                assert = true
                            else
                                assert = lambda( [_Rules].get(role).get(resource).get(privilege).get(1) )(role, resource, privilege)
                            end if
                            if( assert ) then
                                select case [_Rules].get(role).get(resource).get(privilege).get(0)
                                    case "ALLOW":
                                        [_φ] = true
                                    
                                    case "DENY":
                                        [_φ] = false
                                end select
                                exit function
                            end if
                        end if
                    end if
                end if
            next
        next
    end function
    
    ' Function: isAllowed
    ' 
    ' Checks if the role has access to the resource.
    ' 
    ' Parameters:
    ' 
    '     (string) - user identifier
    '     (string) - resource identifier
    '     (string) - privilege identifier
    ' 
    ' Returns:
    ' 
    '     (boolean) - true, if it's allowed; false otherwise
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim AC : set AC = new ACL
    ' set AC.Media = new ACL_Media_MSSQL
    ' AC.Media.connectionString = "Provider=SQLOLEDB;..."
    ' 
    ' call AC.load()
    ' Response.write( AC.isAllowed("nagaozen", "fire-spells", "cast") )
    ' 
    ' set AC = nothing
    ' 
    ' (end code)
    ' 
    public function isAllowed(user, resource, privilege)
        isAllowed = null
        
        dim Sd : set Sd = Server.createObject("Scripting.Dictionary")
        dim this, role, roles, resources
        
        if( not isEmpty( [_Users].get(user) ) ) then
            for each role in [_Users].get(user).keys()
                this = role
                do
                    Sd.add this, null
                    this = [_Roles].get(this).get(0)
                loop while( not isNull(this) )
            next
        end if
        roles = Sd.keys()
        
        call Sd.removeAll()
        
        this = resource
        do
            if( isEmpty( [_Resources].get(this) ) ) then
                this = null
            else
                Sd.add this, null
                this = [_Resources].get(this).get(0)
            end if
        loop while( not isNull(this) )
        resources = Sd.keys()
        
        set Sd = nothing
        
        if( ( not isEmpty(roles) ) and ( not isEmpty(resources) ) ) then
            isAllowed = [_φ](roles, resources, privilege)
        end if
        
        if( isNull(isAllowed) ) then
            isAllowed = false
        end if
    end function
    
    ' Subroutine: load
    ' 
    ' Retrieves the Acl image from the persistence layer.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim AC : set AC = new ACL
    ' set AC.Media = new ACL_Media_MSSQL
    ' AC.Media.connectionString = "Provider=SQLOLEDB;..."
    ' 
    ' call AC.load()
    ' Response.write( AC.isAllowed("nagaozen", "fire-spells", "cast") )
    ' 
    ' set AC = nothing
    ' 
    ' (end code)
    ' 
    public sub load() : call [_ε]
        dim Image : set Image = JSON.parse( Media.load() )
        
        set [_Users] = Image.get("Users")
        set [_Roles] = Image.get("Roles")
        set [_Resources] = Image.get("Resources")
        set [_Rules] = Image.get("Rules")
        
        set Image = nothing
    end sub
    
    ' Subroutine: save
    ' 
    ' Writes the Acl image in the persistence layer.
    ' 
    ' Example:
    ' 
    ' (start code)
    ' 
    ' dim AC : set AC = new ACL
    ' set AC.Media = new ACL_Media_MSSQL
    ' AC.Media.connectionString = "Provider=SQLOLEDB;..."
    ' 
    ' '.
    ' '.Diablo II roles
    ' '.
    ' call AC.addRole("amazon", null)
    ' call AC.addRole("assassin", null)
    ' call AC.addRole("necromancer", null)
    ' call AC.addRole("barbarian", null)
    ' call AC.addRole("paladin", null)
    ' call AC.addRole("sorceress", null)
    ' call AC.addRole("druid", null)
    ' 
    ' '.
    ' '.Sorceress skill trees
    ' '.
    ' call AC.addResource("for-sorceress", null)
    '  
    ' call AC.addResource("cold-spells", "for-sorceress")
    ' call AC.addResource("lightning-spells", "for-sorceress")
    ' call AC.addResource("fire-spells", "for-sorceress")
    ' 
    ' '.
    ' '.Druid skill trees
    ' '.
    ' call AC.addResource("for-druid", null)
    '  
    ' call AC.addResource("elemental", "for-druid")
    ' call AC.addResource("shape-shifting", "for-druid")
    ' call AC.addResource("summoning", "for-druid")
    ' 
    ' call AC.assign("kryfie", "sorceress")
    ' call AC.assign("nagaozen", array("sorceress", "druid"))
    ' 
    ' call AC.allow("sorceress", "cold-spells", "cast", null)
    ' call AC.deny("sorceress", "lightning-spells", "cast", null)
    ' call AC.deny("sorceress", "fire-spells", "cast", null)
    ' 
    ' call AC.save()
    ' 
    ' set AC = nothing
    ' 
    ' (end code)
    ' 
    public sub save() : call [_ε]
        dim Image : set Image = JSON.parse("{}")
        
        call Image.set("Users", [_Users])
        call Image.set("Roles", [_Roles])
        call Image.set("Resources", [_Resources])
        call Image.set("Rules", [_Rules])
        
        call Media.save( JSON.stringify(Image) )
        
        set Image = nothing
    end sub
    
end class

%>
