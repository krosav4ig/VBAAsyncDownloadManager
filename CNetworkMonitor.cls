VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNetworkMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================================
' Класс: CNetworkMonitor
' Назначение: Мониторинг загруженности сети в % от пропускной способности
' ============================================================================
Option Explicit

Private m_maxUsagePercent As Double  ' Максимально допустимая загрузка в % (0 = без лимита)
Private m_currentUsagePercent As Double ' Текущая загрузка в %
Private m_interfaceFilter As String  ' Фильтр имени адаптера (например, "Wi-Fi" или "Ethernet")

' ============================================================================
' Инициализация
' maxUsagePercent: Лимит в процентах (0-100). 0 означает отсутствие лимита.
' interfaceFilter: Часть имени адаптера для фильтрации ("" = суммировать все физические)
' ============================================================================
Public Sub Init(ByVal MaxUsagePercent As Double, Optional ByVal InterfaceFilter As String = "")
    If MaxUsagePercent < 0 Then MaxUsagePercent = 0
    If MaxUsagePercent > 100 Then MaxUsagePercent = 100
    
    m_maxUsagePercent = MaxUsagePercent
    m_interfaceFilter = InterfaceFilter
    m_currentUsagePercent = 0
End Sub

' ============================================================================
' Обновление данных о сети (вызывать из таймера менеджера)
' ============================================================================
Public Sub Refresh()
    On Error GoTo ErrHandler
    
    Dim wmi As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    
    Dim query As String
    ' CurrentBandwidth > 0 отсекает виртуальные адаптеры (VPN, Hyper-V),
    ' у которых часто бывает 0 или некорректная пропускная способность.
    If m_interfaceFilter = "" Then
        query = "SELECT Name, BytesTotalPerSec, CurrentBandwidth FROM Win32_PerfFormattedData_Tcpip_NetworkInterface WHERE CurrentBandwidth > 0"
    Else
        query = "SELECT Name, BytesTotalPerSec, CurrentBandwidth FROM Win32_PerfFormattedData_Tcpip_NetworkInterface WHERE Name LIKE '%" & m_interfaceFilter & "%' AND CurrentBandwidth > 0"
    End If
    
    Dim items As Object
    Set items = wmi.ExecQuery(query)
    
    Dim totalBytesPerSec As Double
    Dim totalBandwidth As Double
    Dim item As Object
    
    For Each item In items
        If Not IsNull(item.BytesTotalPerSec) And Not IsNull(item.CurrentBandwidth) Then
            totalBytesPerSec = totalBytesPerSec + CDbl(item.BytesTotalPerSec)
            totalBandwidth = totalBandwidth + CDbl(item.CurrentBandwidth)
        End If
    Next item
    
    ' Расчет процента загрузки:
    ' BytesTotalPerSec (Байт/сек) * 8 = Бит/сек
    ' (Бит/сек / CurrentBandwidth (Бит/сек)) * 100 = Процент
    If totalBandwidth > 0 Then
        m_currentUsagePercent = ((totalBytesPerSec * 8) / totalBandwidth) * 100
    Else
        m_currentUsagePercent = 0
    End If
    
    ' Округляем до 2 знаков для чистоты логов
    m_currentUsagePercent = Round(m_currentUsagePercent, 2)
    
    Exit Sub

ErrHandler:
    ' В случае ошибки WMI считаем, что лимита нет, чтобы не блокировать загрузки
    m_currentUsagePercent = 0
End Sub

' ============================================================================
' Свойства
' ============================================================================
Public Property Get IsOverLimit() As Boolean
    If m_maxUsagePercent <= 0 Then
        IsOverLimit = False
    Else
        IsOverLimit = (m_currentUsagePercent >= m_maxUsagePercent)
    End If
End Property

Public Property Get CurrentUsagePercent() As Double
    CurrentUsagePercent = m_currentUsagePercent
End Property

Public Property Get MaxUsagePercent() As Double
    MaxUsagePercent = m_maxUsagePercent
End Property

Public Property Get InterfaceFilter() As String
    InterfaceFilter = m_interfaceFilter
End Property
