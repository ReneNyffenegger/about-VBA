'
' vim: ft=basic
'

option explicit

private m_name as string
private m_age  as long

property let name(name_ as string)  ' {
    attribute name.vb_UserMemId = 0
    m_name = name_
end property ' }

property get name() as string ' {
    name = m_name
end property ' }

property let age(age_ as long) ' {
    m_age = age_
end property ' }

property get age() as long ' {
    age = m_age
end property ' }
