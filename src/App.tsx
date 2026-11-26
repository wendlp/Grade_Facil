import { useEffect, useMemo, useRef, useState } from 'react'
import {
  AlertCircle,
  ArrowLeft,
  ArrowRight,
  BookOpen,
  Check,
  ChevronDown,
  Copy,
  Download,
  Edit3,
  FileDown,
  FileUp,
  GraduationCap,
  Plus,
  RefreshCw,
  Search,
  Sparkles,
  Trash2,
  X,
} from 'lucide-react'
import { Document, HeadingLevel, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from 'docx'

type Day = number
type Shift = 'Matutino' | 'Vespertino' | 'Noturno' | 'Integral'
type Subject = {
  id: string
  name: string
  teacher: string
  lessons: number
  days: Day[]
  blocked: string[]
  shifts: Shift[]
  maxClasses: number
  color: number
}
type SchoolClass = {
  id: string
  name: string
  shift: Shift
  lessonsByDay: number[]
  afternoonLessonsByDay: number[]
  lessonDuration: number
  startTime: string
  afternoonStartTime: string
  subjects: Subject[]
}
type ClassDraft = {
  name: string
  shift: Shift
  lessonsByDay: string[]
  afternoonLessonsByDay: string[]
  lessonDuration: string
  startTime: string
  afternoonStartTime: string
}
type Grid = Record<string, string[][]>

const DAYS = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo']
const DAYS_SHORT = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom']
const SHIFTS: Shift[] = ['Matutino', 'Vespertino', 'Noturno', 'Integral']
const COLORS = [
  ['#eff6ff', '#bfdbfe', '#1e40af'],
  ['#fef2f2', '#fecaca', '#991b1b'],
  ['#f0fdf4', '#bbf7d0', '#166534'],
  ['#fefce8', '#fef08a', '#854d0e'],
  ['#faf5ff', '#e9d5ff', '#6b21a8'],
  ['#fff7ed', '#fed7aa', '#9a3412'],
  ['#f0fdfa', '#99f6e4', '#115e59'],
  ['#fdf2f8', '#fbcfe8', '#9d174d'],
  ['#eef2ff', '#c7d2fe', '#3730a3'],
  ['#ecfeff', '#a5f3fc', '#155e75'],
  ['#f7fee7', '#d9f99d', '#3f6212'],
]

const uid = () => Math.random().toString(36).slice(2, 10)
const defaultSubject = (): Omit<Subject, 'id' | 'color'> => ({
  name: '',
  teacher: '',
  lessons: 3,
  days: [0, 1, 2, 3, 4],
  blocked: [],
  shifts: ['Matutino', 'Vespertino', 'Noturno', 'Integral'],
  maxClasses: 1,
})
const defaultClassDraft = (): ClassDraft => ({
  name: '',
  shift: 'Matutino' as Shift,
  lessonsByDay: ['5', '5', '5', '5', '5', '0', '0'],
  afternoonLessonsByDay: ['0', '0', '0', '0', '0', '0', '0'],
  lessonDuration: '50',
  startTime: '07:00',
  afternoonStartTime: '13:00',
})
const normalizeLessonsByDay = (value?: number[], fallback = 5) =>
  Array.from({ length: DAYS.length }, (_, index) => value?.[index] ?? (index < 5 ? fallback : 0))
const classCapacity = (item?: SchoolClass) => item?.lessonsByDay.reduce((sum, count) => sum + count, 0) || 0
const maxLessons = (item?: SchoolClass) => Math.max(...(item?.lessonsByDay || [1]))
const timeToMinutes = (value: string) => {
  const [hours, minutes] = value.split(':').map(Number)
  return hours * 60 + minutes
}
const minutesToTime = (value: number) =>
  `${String(Math.floor((value % 1440) / 60)).padStart(2, '0')}:${String(value % 60).padStart(2, '0')}`
const periodRange = (item: SchoolClass, day: number, period: number) => {
  const morningCount = item.lessonsByDay[day] - (item.afternoonLessonsByDay[day] || 0)
  const isAfternoon = item.shift === 'Integral' && period >= morningCount
  const sessionPeriod = isAfternoon ? period - morningCount : period
  const sessionStart = isAfternoon ? item.afternoonStartTime : item.startTime
  const start = timeToMinutes(sessionStart) + sessionPeriod * item.lessonDuration
  return { start, end: start + item.lessonDuration }
}
const periodLabel = (item: SchoolClass, day: number, period: number) => {
  const range = periodRange(item, day, period)
  return `${minutesToTime(range.start)}–${minutesToTime(range.end)}`
}
const teacherKey = (value: string) =>
  value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, ' ')
    .trim()
    .replace(/\s+/g, ' ')

function findTeacherConflict(classes: SchoolClass[], grid: Grid) {
  const events: {
    teacher: string
    teacherName: string
    className: string
    day: number
    start: number
    end: number
  }[] = []

  for (const schoolClass of classes) {
    for (let day = 0; day < DAYS.length; day++) {
      for (let period = 0; period < schoolClass.lessonsByDay[day]; period++) {
        const subjectId = grid[schoolClass.id]?.[day]?.[period]
        const subject = schoolClass.subjects.find((item) => item.id === subjectId)
        if (!subject) continue
        const range = periodRange(schoolClass, day, period)
        const event = {
          teacher: teacherKey(subject.teacher),
          teacherName: subject.teacher,
          className: schoolClass.name,
          day,
          ...range,
        }
        const conflict = events.find(
          (existing) =>
            existing.teacher === event.teacher &&
            existing.day === event.day &&
            event.start < existing.end &&
            event.end > existing.start,
        )
        if (conflict) return { first: conflict, second: event }
        events.push(event)
      }
    }
  }
  return null
}

function buildSchedule(classes: SchoolClass[]): Grid | null {
  const grid: Grid = {}
  const teacherBusy: { teacher: string; day: number; start: number; end: number }[] = []
  const ordered = [...classes].sort((a, b) => {
    const aOptions = a.subjects.reduce((n, s) => n + s.days.length, 0)
    const bOptions = b.subjects.reduce((n, s) => n + s.days.length, 0)
    return aOptions - bOptions
  })

  const solveClass = (classIndex: number): boolean => {
    if (classIndex === ordered.length) return true
    const current = ordered[classIndex]
    const cells: string[][] = current.lessonsByDay.map((count) => Array(count).fill(''))
    const positions = Array.from({ length: maxLessons(current) }, (_, period) =>
      DAYS.map((_, day) => ({ day: day as Day, period })).filter(({ day }) => period < current.lessonsByDay[day]),
    ).flat()
    const remaining = new Map(current.subjects.map((s) => [s.id, s.lessons]))

    const fill = (position: number): boolean => {
      if (position === positions.length) {
        grid[current.id] = cells
        if (solveClass(classIndex + 1)) return true
        delete grid[current.id]
        return false
      }
      const { day, period } = positions[position]
      const range = periodRange(current, day, period)
      const candidates = current.subjects
        .filter((s) => {
          const key = `${day}-${period}`
          return (
            (remaining.get(s.id) || 0) > 0 &&
            s.days.includes(day) &&
            s.shifts.includes(current.shift) &&
            !s.blocked.includes(key) &&
            !teacherBusy.some(
              (busy) =>
                busy.teacher === teacherKey(s.teacher) &&
                busy.day === day &&
                range.start < busy.end &&
                range.end > busy.start,
            )
          )
        })
        .sort((a, b) => {
          const ar = remaining.get(a.id) || 0
          const br = remaining.get(b.id) || 0
          const aDayCount = cells[day].filter((id) => id === a.id).length
          const bDayCount = cells[day].filter((id) => id === b.id).length
          return aDayCount - bDayCount || br - ar || a.days.length - b.days.length
        })

      for (const candidate of candidates) {
        const key = `${day}-${period}`
        cells[day][period] = candidate.id
        remaining.set(candidate.id, (remaining.get(candidate.id) || 0) - 1)
        teacherBusy.push({ teacher: teacherKey(candidate.teacher), day, ...range })
        if (fill(position + 1)) return true
        teacherBusy.pop()
        remaining.set(candidate.id, (remaining.get(candidate.id) || 0) + 1)
        cells[day][period] = ''
      }
      return false
    }

    return fill(0)
  }

  return solveClass(0) ? grid : null
}

function Button({
  children,
  variant = 'primary',
  className = '',
  ...props
}: React.ButtonHTMLAttributes<HTMLButtonElement> & {
  variant?: 'primary' | 'secondary' | 'outline' | 'danger' | 'ghost'
}) {
  return (
    <button className={`button ${variant} ${className}`} {...props}>
      {children}
    </button>
  )
}

function App() {
  const [step, setStep] = useState(1)
  const [school, setSchool] = useState('')
  const [classes, setClasses] = useState<SchoolClass[]>([])
  const [selectedClassId, setSelectedClassId] = useState('')
  const [classDraft, setClassDraft] = useState(defaultClassDraft())
  const [classModal, setClassModal] = useState<{ mode: 'new' | 'edit'; id?: string; sourceId?: string } | null>(null)
  const [search, setSearch] = useState('')
  const [form, setForm] = useState(defaultSubject())
  const [editingSubjectId, setEditingSubjectId] = useState<string | null>(null)
  const [advanced, setAdvanced] = useState(false)
  const [linkedClassIds, setLinkedClassIds] = useState<string[]>([])
  const [teacherAdvanced, setTeacherAdvanced] = useState(false)
  const [grid, setGrid] = useState<Grid>({})
  const [viewClass, setViewClass] = useState('all')
  const [message, setMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null)
  const [editingCell, setEditingCell] = useState<{ classId: string; day: number; period: number } | null>(null)
  const importRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    const saved = localStorage.getItem('gradefacil-data')
    if (!saved) return
    try {
      const data = JSON.parse(saved)
      setSchool(data.school || school)
      const migrated = (data.classes || []).map((item: Partial<SchoolClass>) => ({
        ...item,
        shift: item.shift || 'Matutino',
        lessonsByDay: normalizeLessonsByDay(item.lessonsByDay, data.lessonsPerDay || 5),
        afternoonLessonsByDay: normalizeLessonsByDay(item.afternoonLessonsByDay, 0),
        lessonDuration: item.lessonDuration || 50,
        startTime: item.startTime || (item.shift === 'Vespertino' ? '13:00' : item.shift === 'Noturno' ? '19:00' : '07:00'),
        afternoonStartTime: item.afternoonStartTime || '13:00',
        subjects: (item.subjects || []).map((s) => ({
          ...s,
          shifts: s.shifts || ['Matutino', 'Vespertino', 'Noturno', 'Integral'],
          maxClasses: s.maxClasses || 1,
        })),
      })) as SchoolClass[]
      setClasses(migrated)
      setSelectedClassId(migrated[0]?.id || '')
      const savedGrid = data.grid || {}
      setGrid(findTeacherConflict(migrated, savedGrid) ? {} : savedGrid)
    } catch {
      localStorage.removeItem('gradefacil-data')
    }
  }, [])

  useEffect(() => {
    localStorage.setItem('gradefacil-data', JSON.stringify({ school, classes, grid }))
  }, [school, classes, grid])

  useEffect(() => {
    if (!selectedClassId && classes[0]) setSelectedClassId(classes[0].id)
    if (selectedClassId && !classes.some((c) => c.id === selectedClassId)) setSelectedClassId(classes[0]?.id || '')
  }, [classes, selectedClassId])

  const selectedClass = classes.find((c) => c.id === selectedClassId)
  const capacity = classCapacity(selectedClass)
  const currentTotal = selectedClass?.subjects.reduce((sum, s) => sum + s.lessons, 0) || 0
  const visibleClasses = classes.filter((c) => c.name.toLowerCase().includes(search.toLowerCase()))
  const teacherRegistrations = useMemo(() => {
    const registrations = new Map<string, { teacher: string; subjects: Set<string>; classes: Set<string>; maxClasses: number }>()
    classes.forEach((schoolClass) =>
      schoolClass.subjects.forEach((subjectItem) => {
        const key = teacherKey(subjectItem.teacher)
        const current = registrations.get(key) || {
          teacher: subjectItem.teacher,
          subjects: new Set<string>(),
          classes: new Set<string>(),
          maxClasses: subjectItem.maxClasses,
        }
        current.subjects.add(subjectItem.name)
        current.classes.add(schoolClass.name)
        current.maxClasses = subjectItem.maxClasses
        registrations.set(key, current)
      }),
    )
    return [...registrations.values()]
  }, [classes])

  const notify = (text: string, type: 'success' | 'error' = 'success') => {
    setMessage({ text, type })
    window.setTimeout(() => setMessage(null), 3500)
  }

  const addTeacherRegistration = () => {
    if (!form.teacher.trim() || !form.name.trim()) return notify('Informe o professor e a disciplina.', 'error')
    if (!linkedClassIds.length) return notify('Selecione ao menos uma turma vinculada.', 'error')
    if (!form.days.length || !form.shifts.length) return notify('Selecione os dias e turnos disponíveis.', 'error')
    const selectedClasses = classes.filter((item) => linkedClassIds.includes(item.id))
    const unavailableShift = selectedClasses.find((item) => !form.shifts.includes(item.shift))
    if (unavailableShift)
      return notify(`O professor precisa estar disponível no turno ${unavailableShift.shift.toLowerCase()} de ${unavailableShift.name}.`, 'error')
    if (form.lessons < 1 || form.lessons > 100) return notify('Informe uma quantidade válida de aulas por turma.', 'error')
    const existingTeacherClasses = new Set<string>()
    classes.forEach((schoolClass) => {
      if (schoolClass.subjects.some((item) => teacherKey(item.teacher) === teacherKey(form.teacher)))
        existingTeacherClasses.add(schoolClass.name.toLowerCase())
    })
    selectedClasses.forEach((item) => existingTeacherClasses.add(item.name.toLowerCase()))
    if (form.maxClasses < existingTeacherClasses.size)
      return notify(
        `${form.teacher} ficará em ${existingTeacherClasses.size} turmas. Aumente o limite informado.`,
        'error',
      )
    const duplicate = selectedClasses.find((existing) => {
      return existing.subjects.some(
        (item) => item.name.toLowerCase() === form.name.trim().toLowerCase() &&
          teacherKey(item.teacher) === teacherKey(form.teacher),
      )
    })
    if (duplicate) return notify(`Este professor e disciplina já estão vinculados à turma ${duplicate.name}.`, 'error')

    let next = classes.map((item) => ({
      ...item,
      subjects: item.subjects.map((subjectItem) =>
        teacherKey(subjectItem.teacher) === teacherKey(form.teacher)
          ? { ...subjectItem, maxClasses: form.maxClasses }
          : subjectItem,
      ),
    }))
    for (const selectedClass of selectedClasses) {
      const schoolClass = next.find((item) => item.id === selectedClass.id)!
      schoolClass.subjects.push({
        ...form,
        id: uid(),
        name: form.name.trim(),
        teacher: form.teacher.trim(),
        color: schoolClass.subjects.length % COLORS.length,
      })
    }
    setClasses(next)
    setSelectedClassId((current) => current || next[0]?.id || '')
    setGrid({})
    setForm(defaultSubject())
    setLinkedClassIds([])
    setTeacherAdvanced(false)
    notify('Professor, disciplina e turmas vinculados com sucesso.')
  }

  const openNewClass = () => {
    setClassDraft(defaultClassDraft())
    setClassModal({ mode: 'new' })
  }

  const saveClass = () => {
    const name = classDraft.name.trim()
    if (!name) return notify('Informe o nome da turma.', 'error')
    if (classes.some((c) => c.id !== classModal?.id && c.name.toLowerCase() === name.toLowerCase()))
      return notify('Já existe uma turma com esse nome.', 'error')
    const primaryLessons = classDraft.lessonsByDay.map(Number)
    const afternoonLessons = classDraft.afternoonLessonsByDay.map(Number)
    if ([...primaryLessons, ...afternoonLessons].some((count) => !Number.isFinite(count) || count < 0 || count > 12))
      return notify('Cada dia deve ter entre 0 e 12 aulas.', 'error')
    if (!primaryLessons.some((count) => count > 0) && !(classDraft.shift === 'Integral' && afternoonLessons.some((count) => count > 0)))
      return notify('Defina ao menos um dia com aulas.', 'error')
    const lessonDuration = Number(classDraft.lessonDuration)
    if (!Number.isFinite(lessonDuration) || lessonDuration < 10 || lessonDuration > 240)
      return notify('A duração deve estar entre 10 e 240 minutos.', 'error')
    if (!/^\d{2}:\d{2}$/.test(classDraft.startTime))
      return notify('Informe um horário inicial válido.', 'error')
    if (classDraft.shift === 'Integral' && !/^\d{2}:\d{2}$/.test(classDraft.afternoonStartTime))
      return notify('Informe um horário inicial válido para a tarde.', 'error')
    if (
      classDraft.shift === 'Integral' &&
      primaryLessons.some(
        (count, day) =>
          count > 0 &&
          afternoonLessons[day] > 0 &&
          timeToMinutes(classDraft.afternoonStartTime) < timeToMinutes(classDraft.startTime) + count * lessonDuration,
      )
    )
      return notify('O início da tarde não pode coincidir com as aulas da manhã.', 'error')
    const savedAfternoon = classDraft.shift === 'Integral' ? afternoonLessons : Array(DAYS.length).fill(0)
    const savedLessonsByDay = primaryLessons.map((count, index) => count + savedAfternoon[index])
    const nextCapacity = savedLessonsByDay.reduce((sum, count) => sum + count, 0)
    if (classModal?.mode === 'edit' && classModal.id) {
      const current = classes.find((c) => c.id === classModal.id)
      const assigned = current?.subjects.reduce((sum, s) => sum + s.lessons, 0) || 0
      if (assigned > nextCapacity)
        return notify(`A nova capacidade é menor que as ${assigned} aulas já cadastradas.`, 'error')
      setClasses((items) =>
        items.map((c) =>
          c.id === classModal.id
            ? { ...c, name, shift: classDraft.shift, lessonsByDay: savedLessonsByDay, afternoonLessonsByDay: savedAfternoon, lessonDuration, startTime: classDraft.startTime, afternoonStartTime: classDraft.afternoonStartTime }
            : c,
        ),
      )
      setGrid({})
      notify('Configurações da turma atualizadas.')
    } else {
      const source = classes.find((item) => item.id === classModal?.sourceId)
      const next: SchoolClass = {
        id: uid(),
        name,
        shift: classDraft.shift,
        lessonsByDay: savedLessonsByDay,
        afternoonLessonsByDay: savedAfternoon,
        lessonDuration,
        startTime: classDraft.startTime,
        afternoonStartTime: classDraft.afternoonStartTime,
        subjects: source?.subjects.map((item) => ({ ...item, id: uid() })) || [],
      }
      setClasses((items) => [...items, next])
      setSelectedClassId(next.id)
      setGrid({})
      notify(source ? 'Turma replicada. Revise os vínculos dos professores.' : 'Turma adicionada.')
    }
    setClassModal(null)
  }

  const editClass = (item: SchoolClass) => {
    setClassDraft({
      name: item.name,
      shift: item.shift,
      lessonsByDay: item.lessonsByDay.map((count, index) => String(count - (item.afternoonLessonsByDay[index] || 0))),
      afternoonLessonsByDay: item.afternoonLessonsByDay.map(String),
      lessonDuration: String(item.lessonDuration),
      startTime: item.startTime,
      afternoonStartTime: item.afternoonStartTime,
    })
    setClassModal({ mode: 'edit', id: item.id })
  }

  const duplicateClass = (item: SchoolClass) => {
    setClassDraft({
      name: `${item.name} - Cópia`,
      shift: item.shift,
      lessonsByDay: item.lessonsByDay.map((count, index) => String(count - (item.afternoonLessonsByDay[index] || 0))),
      afternoonLessonsByDay: item.afternoonLessonsByDay.map(String),
      lessonDuration: String(item.lessonDuration),
      startTime: item.startTime,
      afternoonStartTime: item.afternoonStartTime,
    })
    setClassModal({ mode: 'new', sourceId: item.id })
  }

  const deleteClass = (id: string) => {
    if (!window.confirm('Excluir esta turma e todas as disciplinas cadastradas?')) return
    setClasses((items) => items.filter((c) => c.id !== id))
    setGrid((old) => {
      const next = { ...old }
      delete next[id]
      return next
    })
  }

  const saveSubject = () => {
    if (!selectedClass) return
    if (!form.name.trim() || !form.teacher.trim()) return notify('Preencha a disciplina e o professor.', 'error')
    if (!form.days.length) return notify('Selecione ao menos um dia disponível.', 'error')
    if (!form.shifts.length) return notify('Selecione ao menos um turno disponível para o professor.', 'error')
    if (!Number.isInteger(form.maxClasses) || form.maxClasses < 1 || form.maxClasses > 100)
      return notify('Informe uma quantidade válida de turmas para o professor.', 'error')
    if (form.lessons < 1 || form.lessons > capacity) return notify('A carga horária informada é inválida.', 'error')
    const others = selectedClass.subjects.filter((s) => s.id !== editingSubjectId)
    const total = others.reduce((n, s) => n + s.lessons, 0) + form.lessons
    if (total > capacity) return notify(`A carga excede a capacidade de ${capacity} aulas.`, 'error')
    const updated: Subject = {
      ...form,
      id: editingSubjectId || uid(),
      name: form.name.trim(),
      teacher: form.teacher.trim(),
      color: editingSubjectId
        ? selectedClass.subjects.find((s) => s.id === editingSubjectId)?.color || 0
        : selectedClass.subjects.length % COLORS.length,
    }
    setClasses((items) =>
      items.map((c) => {
        const subjects =
          c.id === selectedClass.id
            ? editingSubjectId
              ? c.subjects.map((s) => (s.id === editingSubjectId ? updated : s))
              : [...c.subjects, updated]
            : c.subjects
        return {
          ...c,
          subjects: subjects.map((s) =>
            teacherKey(s.teacher) === teacherKey(updated.teacher) ? { ...s, maxClasses: updated.maxClasses } : s,
          ),
        }
      }),
    )
    setForm(defaultSubject())
    setEditingSubjectId(null)
    setAdvanced(false)
    setGrid({})
    notify(editingSubjectId ? 'Disciplina atualizada.' : 'Disciplina adicionada.')
  }

  const editSubject = (item: Subject) => {
    setForm({ name: item.name, teacher: item.teacher, lessons: item.lessons, days: item.days, blocked: item.blocked, shifts: item.shifts, maxClasses: item.maxClasses })
    setEditingSubjectId(item.id)
    setAdvanced(item.blocked.length > 0)
    window.scrollTo({ top: 250, behavior: 'smooth' })
  }

  const removeSubject = (id: string) => {
    if (!selectedClass || !window.confirm('Excluir esta disciplina?')) return
    setClasses((items) =>
      items.map((c) => (c.id === selectedClass.id ? { ...c, subjects: c.subjects.filter((s) => s.id !== id) } : c)),
    )
    setGrid({})
  }

  const goToSubjects = (classId?: string) => {
    if (!school.trim()) return notify('Informe o nome da escola.', 'error')
    if (!classes.length) return notify('Adicione ao menos uma turma.', 'error')
    if (classId) setSelectedClassId(classId)
    setStep(2)
    window.scrollTo(0, 0)
  }

  const generate = () => {
    const teacherAssignments = new Map<string, { name: string; classes: Set<string>; limit: number; declaredLimits: Set<number> }>()
    classes.forEach((schoolClass) =>
      schoolClass.subjects.forEach((subjectItem) => {
        const key = teacherKey(subjectItem.teacher)
        const current = teacherAssignments.get(key) || {
          name: subjectItem.teacher,
          classes: new Set<string>(),
          limit: subjectItem.maxClasses,
          declaredLimits: new Set<number>(),
        }
        current.classes.add(schoolClass.id)
        current.declaredLimits.add(subjectItem.maxClasses)
        current.limit = Math.min(current.limit, subjectItem.maxClasses)
        teacherAssignments.set(key, current)
      }),
    )
    const inconsistentTeacher = [...teacherAssignments.values()].find((teacher) => teacher.declaredLimits.size > 1)
    if (inconsistentTeacher)
      return notify(
        `${inconsistentTeacher.name} possui quantidades de turmas diferentes nos cadastros. Use o mesmo limite em todas as disciplinas.`,
        'error',
      )
    const overloadedTeacher = [...teacherAssignments.values()].find((teacher) => teacher.classes.size > teacher.limit)
    if (overloadedTeacher)
      return notify(
        `${overloadedTeacher.name} está em ${overloadedTeacher.classes.size} turmas, mas o limite informado é ${overloadedTeacher.limit}.`,
        'error',
      )
    const unavailable = classes.find((c) => c.subjects.some((s) => !s.shifts.includes(c.shift)))
    if (unavailable)
      return notify(`${unavailable.name} possui professor sem disponibilidade no turno ${unavailable.shift.toLowerCase()}.`, 'error')
    const incomplete = classes.find((c) => c.subjects.reduce((sum, s) => sum + s.lessons, 0) !== classCapacity(c))
    if (incomplete)
      return notify(`${incomplete.name} precisa ter exatamente ${classCapacity(incomplete)} aulas cadastradas.`, 'error')
    const result = buildSchedule(classes)
    if (!result)
      return notify('Não foi possível conciliar as disponibilidades. Libere mais dias ou horários e tente novamente.', 'error')
    const conflict = findTeacherConflict(classes, result)
    if (conflict)
      return notify(
        `Conflito detectado: ${conflict.second.teacherName} em ${conflict.first.className} e ${conflict.second.className}.`,
        'error',
      )
    setGrid(result)
    setStep(3)
    setViewClass('all')
    window.scrollTo(0, 0)
    notify('Grade gerada sem conflitos.')
  }

  const manualChange = (subjectId: string) => {
    if (!editingCell) return
    const { classId, day, period } = editingCell
    const targetClass = classes.find((c) => c.id === classId)
    const nextSubject = targetClass?.subjects.find((s) => s.id === subjectId)
    if (!targetClass || !nextSubject) return
    const currentSubjectId = grid[classId]?.[day]?.[period]
    if (subjectId === currentSubjectId) {
      setEditingCell(null)
      return
    }
    const currentSubject = targetClass.subjects.find((s) => s.id === currentSubjectId)
    if (!currentSubject) return
    if (!nextSubject.days.includes(day as Day) || nextSubject.blocked.includes(`${day}-${period}`))
      return notify('O professor não está disponível nesse horário.', 'error')
    const teacherHasConflict = (teacher: string, checkDay: number, checkPeriod: number) => {
      const targetRange = periodRange(targetClass, checkDay, checkPeriod)
      return (
      classes.some((c) => {
        if (c.id === classId) return false
        return (grid[c.id]?.[checkDay] || []).some((otherId, otherPeriod) => {
          const other = c.subjects.find((s) => s.id === otherId)
          if (teacherKey(other?.teacher || '') !== teacherKey(teacher)) return false
          const otherRange = periodRange(c, checkDay, otherPeriod)
          return targetRange.start < otherRange.end && targetRange.end > otherRange.start
        })
      })
      )
    }
    if (teacherHasConflict(nextSubject.teacher, day, period))
      return notify('Esse professor já está em outra turma nesse horário.', 'error')

    let swap: { day: number; period: number } | null = null
      for (let candidateDay = 0; candidateDay < DAYS.length && !swap; candidateDay++) {
      for (let candidatePeriod = 0; candidatePeriod < targetClass.lessonsByDay[candidateDay]; candidatePeriod++) {
        if (grid[classId]?.[candidateDay]?.[candidatePeriod] !== subjectId) continue
        const key = `${candidateDay}-${candidatePeriod}`
        if (
          currentSubject.days.includes(candidateDay as Day) &&
          !currentSubject.blocked.includes(key) &&
          !teacherHasConflict(currentSubject.teacher, candidateDay, candidatePeriod)
        ) {
          swap = { day: candidateDay, period: candidatePeriod }
          break
        }
      }
    }
    if (!swap)
      return notify('Não há outra aula compatível para realizar a troca sem alterar a carga horária.', 'error')

    const nextGrid: Grid = {
      ...grid,
      [classId]: grid[classId].map((col, d) =>
        col.map((id, p) => {
          if (d === day && p === period) return subjectId
          if (d === swap.day && p === swap.period) return currentSubjectId
          return id
        }),
      ),
    }
    const auditedConflict = findTeacherConflict(classes, nextGrid)
    if (auditedConflict)
      return notify(
        `Troca bloqueada: ${auditedConflict.second.teacherName} ficaria em duas turmas ao mesmo tempo.`,
        'error',
      )
    setGrid(nextGrid)
    setEditingCell(null)
    notify('Aulas trocadas sem alterar as cargas horárias.')
  }

  const exportWord = async () => {
    const marker = btoa(unescape(encodeURIComponent(JSON.stringify({ school, classes, grid }))))
    const children: (Paragraph | Table)[] = [
      new Paragraph({ text: school, heading: HeadingLevel.TITLE }),
      new Paragraph({ text: 'Grade horária escolar', heading: HeadingLevel.HEADING_1 }),
      new Paragraph({ children: [new TextRun({ text: `GRADE_FACIL_DATA:${marker}`, color: 'FFFFFF', size: 2 })] }),
    ]
    classes.forEach((c) => {
      children.push(new Paragraph({ text: c.name, heading: HeadingLevel.HEADING_2 }))
      children.push(new Paragraph({ text: `Turno: ${c.shift} · ${classCapacity(c)} aulas por semana · ${c.lessonDuration} minutos por aula · ${c.shift === 'Integral' ? `manhã às ${c.startTime} e tarde às ${c.afternoonStartTime}` : `início às ${c.startTime}`}` }))
      children.push(
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph('Horário')] }),
                ...DAYS.map((d) => new TableCell({ children: [new Paragraph(d)] })),
              ],
            }),
            ...Array.from({ length: maxLessons(c) }, (_, period) =>
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph(`${period + 1}ª aula`)] }),
                  ...DAYS.map((_, day) => {
                    const s = c.subjects.find((x) => x.id === grid[c.id]?.[day]?.[period])
                    const value = period >= c.lessonsByDay[day] ? 'Sem aula' : s ? `${periodLabel(c, day, period)}\n${s.name}\n${s.teacher}` : '—'
                    return new TableCell({ children: [new Paragraph(value)] })
                  }),
                ],
              }),
            ),
          ],
        }),
      )
    })
    const blob = await Packer.toBlob(new Document({ sections: [{ children }] }))
    downloadBlob(blob, `grade-${safeName(school)}.docx`)
    notify('Arquivo Word exportado.')
  }

  const importWord = async (file: File) => {
    try {
      const mammoth = await import('mammoth/mammoth.browser')
      const buffer = await file.arrayBuffer()
      const result = await mammoth.extractRawText({ arrayBuffer: buffer })
      const match = result.value.match(/GRADE_FACIL_DATA:([A-Za-z0-9+/=]+)/)
      if (!match) throw new Error('Arquivo incompatível')
      const data = JSON.parse(decodeURIComponent(escape(atob(match[1]))))
      if (!Array.isArray(data.classes)) throw new Error('Dados inválidos')
      setSchool(data.school)
      const imported = data.classes.map((item: SchoolClass) => ({
        ...item,
        shift: item.shift || 'Matutino',
        lessonsByDay: normalizeLessonsByDay(item.lessonsByDay, data.lessonsPerDay || 5),
        afternoonLessonsByDay: normalizeLessonsByDay(item.afternoonLessonsByDay, 0),
        lessonDuration: item.lessonDuration || 50,
        startTime: item.startTime || (item.shift === 'Vespertino' ? '13:00' : item.shift === 'Noturno' ? '19:00' : '07:00'),
        afternoonStartTime: item.afternoonStartTime || '13:00',
        subjects: item.subjects.map((s) => ({ ...s, shifts: s.shifts || SHIFTS, maxClasses: s.maxClasses || 1 })),
      }))
      setClasses(imported)
      const importedGrid = data.grid || {}
      setGrid(findTeacherConflict(imported, importedGrid) ? {} : importedGrid)
      setSelectedClassId(imported[0]?.id || '')
      setStep(1)
      notify('Arquivo importado com sucesso.')
    } catch {
      notify('Não foi possível importar. Use um arquivo exportado pelo GradeFácil.', 'error')
    }
  }

  const classesToShow = classes.filter((c) => viewClass === 'all' || c.id === viewClass)

  return (
    <div className="app-shell">
      <header className="header">
        <div className="topbar">
          <button className="brand" onClick={() => setStep(1)} aria-label="Voltar ao início">
            <img className="brand-logo" src="/logo.png" alt="" width={44} height={44} />
            <span>
              <span className="brand-title">GradeFácil <strong>Escolar</strong></span>
              <span className="brand-subtitle">Horários escolares sem conflitos</span>
            </span>
          </button>
          <div className="header-actions">
            <span className="offline-badge"><span /> Dados salvos neste dispositivo</span>
            <Button variant="outline" onClick={() => importRef.current?.click()}><FileUp size={16} /> Importar .docx</Button>
            <input
              ref={importRef}
              hidden
              type="file"
              accept=".docx"
              onChange={(e) => e.target.files?.[0] && importWord(e.target.files[0])}
            />
          </div>
        </div>
        <nav className="stepper" aria-label="Etapas">
          {['Cadastrar turmas', 'Professores e aulas', 'Grade completa'].map((label, i) => (
            <button key={label} className={step === i + 1 ? 'step active' : step > i + 1 ? 'step done' : 'step'} onClick={() => step > i + 1 && setStep(i + 1)}>
              <span>{step > i + 1 ? <Check size={14} /> : i + 1}</span>
              <small>ETAPA {i + 1}</small>
              <strong>{label}</strong>
            </button>
          ))}
        </nav>
      </header>

      <main>
        {message && <div className={`toast ${message.type}`}><AlertCircle size={18} />{message.text}<button onClick={() => setMessage(null)}><X size={16} /></button></div>}
        {step === 1 && (
          <>
            <section className="hero" aria-labelledby="hero-title">
              <img className="hero-logo" src="/logo.png" alt="Logotipo GradeFácil Escolar" width={132} height={132} />
              <div className="hero-text">
                <span className="hero-eyebrow">Planejamento escolar</span>
                <h1 id="hero-title">GradeFácil <strong>Escolar</strong></h1>
                <p>Cadastre turmas e professores, gere horários sem conflitos e exporte a grade completa para PDF ou Word.</p>
                <a className="hero-cta" href="#cadastro">Começar agora <ArrowRight size={17} /></a>
              </div>
            </section>
            <div className="page-heading" id="cadastro"><div><span className="eyebrow">PRIMEIRO PASSO</span><h2>Cadastre todas as turmas</h2><p>Configure turno, carga diária, duração e horários antes de vincular professores.</p></div></div>
            <div className="two-columns structure-grid">
              <section className="card">
                <CardHeading icon={<BookOpen size={19} />} title="Dados da escola" subtitle="Informações gerais da grade" />
                <Field label="Nome da escola"><input value={school} onChange={(e) => setSchool(e.target.value)} placeholder="Nome da instituição" /></Field>
                <div className="info-panel"><GraduationCap size={18} /><div><strong>Ordem obrigatória</strong><p>Cadastre e configure as turmas primeiro. Somente turmas existentes poderão ser selecionadas no cadastro dos professores.</p></div></div>
              </section>
              <section className="card classes-card">
                <CardHeading icon={<GraduationCap size={19} />} title="Turmas cadastradas" subtitle={`${classes.length} ${classes.length === 1 ? 'turma cadastrada' : 'turmas cadastradas'}`} />
                <Button className="new-class-button" onClick={openNewClass}><Plus size={17} /> Criar nova turma</Button>
                <div className="search"><Search size={17} /><input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Buscar turma..." /></div>
                <div className="class-list">
                  {!visibleClasses.length && <Empty text={classes.length ? 'Nenhuma turma encontrada.' : 'Cadastre sua primeira turma para começar.'} />}
                  {visibleClasses.map((item) => <div className="class-row" key={item.id}><span className="class-icon"><GraduationCap size={19} /></span><div><strong>{item.name} <em>{item.shift}</em></strong><small>{classCapacity(item)} aulas semanais · início {item.startTime}</small></div><button className="icon-button" onClick={() => duplicateClass(item)} aria-label={`Replicar ${item.name}`} title="Replicar turma"><Copy size={16} /></button><button className="icon-button" onClick={() => editClass(item)} aria-label="Editar turma"><Edit3 size={16} /></button><button className="icon-button danger-icon" onClick={() => deleteClass(item.id)} aria-label="Excluir turma"><Trash2 size={16} /></button></div>)}
                </div>
              </section>
            </div>
            <div className="bottom-nav"><span /><Button onClick={() => goToSubjects()}>Próximo: professores <ArrowRight size={17} /></Button></div>
          </>
        )}

        {step === 2 && (
          <>
            <div className="page-heading"><div><span className="eyebrow">SEGUNDO PASSO</span><h1>Cadastre professores e aulas</h1><p>Selecione uma ou mais turmas já cadastradas e informe toda a disponibilidade do professor.</p></div></div>
            <div className="two-columns teacher-flow-grid">
              <section className="card teacher-form-card">
                <CardHeading icon={<Plus size={19} />} title="Professor e disciplina" subtitle="Vincule o cadastro somente às turmas existentes" />
                <div className="form-grid"><Field label="Professor(a)"><input value={form.teacher} onChange={(e) => setForm({ ...form, teacher: e.target.value })} placeholder="Ex.: Prof. Carlos" /></Field><Field label="Disciplina"><input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="Ex.: Matemática" /></Field></div>
                <Field label="Selecionar turmas"><div className="class-selection">{classes.map((item) => <button className={linkedClassIds.includes(item.id) ? 'active' : ''} key={item.id} onClick={() => setLinkedClassIds(linkedClassIds.includes(item.id) ? linkedClassIds.filter((id) => id !== item.id) : [...linkedClassIds, item.id])}>{linkedClassIds.includes(item.id) && <Check size={14} />}<span>{item.name}<small>{item.shift}</small></span></button>)}</div></Field>
                <div className="form-grid"><Field label="Aulas por semana em cada turma"><input type="number" min="1" value={form.lessons || ''} onChange={(e) => setForm({ ...form, lessons: Number(e.target.value) })} /></Field><Field label="Máximo de turmas do professor"><input type="number" min="1" max="100" value={form.maxClasses || ''} onChange={(e) => setForm({ ...form, maxClasses: Number(e.target.value) })} /></Field></div>
                <Field label="Turnos disponíveis do professor"><div className="shift-pills">{SHIFTS.map((shift) => <button className={form.shifts.includes(shift) ? 'active' : ''} key={shift} onClick={() => setForm({ ...form, shifts: form.shifts.includes(shift) ? form.shifts.filter((x) => x !== shift) : [...form.shifts, shift] })}>{form.shifts.includes(shift) && <Check size={13} />}{shift}</button>)}</div></Field>
                <Field label="Dias disponíveis"><div className="day-pills interactive">{DAYS_SHORT.map((day, index) => <button className={form.days.includes(index) ? 'active' : ''} key={day} onClick={() => setForm({ ...form, days: form.days.includes(index) ? form.days.filter((x) => x !== index) : [...form.days, index] })}>{form.days.includes(index) && <Check size={13} />}{day}</button>)}</div></Field>
                <button className="advanced-toggle" onClick={() => setTeacherAdvanced(!teacherAdvanced)}><ChevronDown size={16} className={teacherAdvanced ? 'rotated' : ''} /> Definir momentos indisponíveis</button>
                {teacherAdvanced && <BlockedGrid value={form.blocked} onChange={(blocked) => setForm({ ...form, blocked })} lessons={12} />}
                <Button className="grow register-teacher-button" onClick={addTeacherRegistration}><Plus size={17} /> Adicionar à distribuição</Button>
              </section>
              <section className="card teacher-register-list">
                <CardHeading icon={<GraduationCap size={19} />} title="Distribuição cadastrada" subtitle={`${teacherRegistrations.length} ${teacherRegistrations.length === 1 ? 'professor' : 'professores'} nas ${classes.length} turmas`} />
                {!teacherRegistrations.length && <Empty text="Cadastre o primeiro professor para começar a distribuição." />}
                <div className="teacher-cards">{teacherRegistrations.map((registration) => <div className="teacher-card" key={teacherKey(registration.teacher)}><span className="class-icon"><BookOpen size={18} /></span><div><strong>{registration.teacher}</strong><small>{[...registration.subjects].join(', ')}</small><p>{[...registration.classes].join(' · ')}</p></div><em>{registration.classes.size}/{registration.maxClasses} turmas</em></div>)}</div>
                <div className="class-load-summary">{classes.map((item) => { const assigned = item.subjects.reduce((sum, subjectItem) => sum + subjectItem.lessons, 0); return <div key={item.id}><span><strong>{item.name}</strong><small>{assigned}/{classCapacity(item)} aulas</small></span><div className="progress"><span style={{ width: `${Math.min(100, assigned / classCapacity(item) * 100)}%` }} /></div></div> })}</div>
              </section>
            </div>
            <div className="bottom-nav"><Button variant="outline" onClick={() => { setStep(1); window.scrollTo(0, 0) }}><ArrowLeft size={17} /> Voltar às turmas</Button><Button onClick={generate}><Sparkles size={17} /> Gerar grade completa</Button></div>
          </>
        )}

        {step === 3 && (
          <>
            <div className="page-heading grade-heading"><div><span className="eyebrow">ETAPA 3</span><h1>Grade horária gerada</h1><p>Selecione uma aula para fazer ajustes manuais.</p></div><div className="grade-actions"><Button variant="secondary" onClick={generate}><RefreshCw size={16} /> Gerar novamente</Button><Button variant="outline" onClick={() => window.print()}><FileDown size={16} /> Exportar PDF</Button><Button onClick={exportWord}><Download size={16} /> Exportar Word</Button></div></div>
            <div className="toolbar"><div><label htmlFor="class-filter">Exibir turma</label><select id="class-filter" value={viewClass} onChange={(e) => setViewClass(e.target.value)}><option value="all">Todas as turmas ({classes.length})</option>{classes.map((c) => <option value={c.id} key={c.id}>{c.name}</option>)}</select></div><span className="success-banner"><Check size={17} /> Auditoria concluída: nenhum professor está em duas turmas no mesmo horário.</span></div>
            <div className="schedules">
              {classesToShow.map((c) => <Schedule key={c.id} item={c} grid={grid[c.id]} onCell={(day, period) => setEditingCell({ classId: c.id, day, period })} />)}
            </div>
            <div className="bottom-nav"><Button variant="outline" onClick={() => { setStep(2); window.scrollTo(0, 0) }}><ArrowLeft size={17} /> Voltar às disciplinas</Button><span /></div>
          </>
        )}
      </main>

      {classModal && (
        <div className="modal-backdrop" onMouseDown={() => setClassModal(null)}>
          <div className="modal class-modal" onMouseDown={(e) => e.stopPropagation()}>
            <div className="modal-head">
              <div><span className="eyebrow">{classModal.sourceId ? 'REPLICAR TURMA' : classModal.mode === 'new' ? 'NOVA TURMA' : 'EDITAR TURMA'}</span><h2>{classModal.sourceId ? 'Revise o nome da cópia' : classModal.mode === 'new' ? 'Configure a turma' : 'Dados e carga horária'}</h2></div>
              <button className="icon-button" onClick={() => setClassModal(null)} aria-label="Fechar"><X size={19} /></button>
            </div>
            <p>{classModal.sourceId ? 'A estrutura, os horários e as disciplinas foram copiados. Altere o nome e salve a nova turma.' : 'Defina o turno e quantas aulas essa turma terá em cada dia da semana.'}</p>
            <Field label="Nome da turma"><input autoFocus value={classDraft.name} onChange={(e) => setClassDraft({ ...classDraft, name: e.target.value })} placeholder="Ex.: 8º Ano B" /></Field>
            <Field label="Turno da turma">
              <div className="shift-selector">{SHIFTS.map((shift) => <button className={classDraft.shift === shift ? 'active' : ''} onClick={() => setClassDraft({ ...classDraft, shift, startTime: shift === 'Vespertino' ? '13:00' : shift === 'Noturno' ? '19:00' : classDraft.startTime })} key={shift}>{classDraft.shift === shift && <Check size={14} />}{shift}</button>)}</div>
            </Field>
            <Field label={classDraft.shift === 'Integral' ? 'Aulas da manhã por dia' : `Aulas por dia — ${classDraft.shift.toLowerCase()}`}>
              <div className="daily-lessons">
                {DAYS_SHORT.map((day, index) => <label key={day}><strong>{day}</strong><input type="number" min="0" max="12" value={classDraft.lessonsByDay[index]} onChange={(e) => setClassDraft({ ...classDraft, lessonsByDay: classDraft.lessonsByDay.map((n, i) => i === index ? e.target.value : n) })} /></label>)}
              </div>
            </Field>
            {classDraft.shift === 'Integral' && (
              <Field label="Aulas da tarde por dia">
                <div className="daily-lessons afternoon-counts">
                  {DAYS_SHORT.map((day, index) => <label key={day}><strong>{day}</strong><input type="number" min="0" max="12" value={classDraft.afternoonLessonsByDay[index]} onChange={(e) => setClassDraft({ ...classDraft, afternoonLessonsByDay: classDraft.afternoonLessonsByDay.map((n, i) => i === index ? e.target.value : n) })} /></label>)}
                </div>
              </Field>
            )}
            <div className="time-settings">
              <Field label="Duração de cada aula (minutos)"><input type="number" min="10" max="240" value={classDraft.lessonDuration} onChange={(e) => setClassDraft({ ...classDraft, lessonDuration: e.target.value })} /></Field>
              <Field label={`Início das aulas ${classDraft.shift === 'Vespertino' ? 'da tarde' : classDraft.shift === 'Noturno' ? 'da noite' : 'da manhã'}`}><input type="time" value={classDraft.startTime} onChange={(e) => setClassDraft({ ...classDraft, startTime: e.target.value })} /></Field>
              {classDraft.shift === 'Integral' && <Field label="Início das aulas da tarde"><input type="time" value={classDraft.afternoonStartTime} onChange={(e) => setClassDraft({ ...classDraft, afternoonStartTime: e.target.value })} /></Field>}
            </div>
            <div className="capacity modal-capacity"><span>Capacidade semanal</span><strong>{classDraft.lessonsByDay.reduce((sum, n) => sum + (Number(n) || 0), 0) + (classDraft.shift === 'Integral' ? classDraft.afternoonLessonsByDay.reduce((sum, n) => sum + (Number(n) || 0), 0) : 0)} aulas</strong></div>
            <div className="modal-footer"><Button variant="ghost" onClick={() => setClassModal(null)}>Cancelar</Button><Button onClick={saveClass}><Check size={16} /> {classModal.sourceId ? 'Criar cópia' : classModal.mode === 'new' ? 'Criar turma' : 'Salvar alterações'}</Button></div>
          </div>
        </div>
      )}

      {editingCell && (() => {
        const c = classes.find((x) => x.id === editingCell.classId)!
        const currentId = grid[c.id]?.[editingCell.day]?.[editingCell.period]
        return <div className="modal-backdrop" onMouseDown={() => setEditingCell(null)}><div className="modal" onMouseDown={(e) => e.stopPropagation()}><div className="modal-head"><div><span className="eyebrow">ALTERAR AULA</span><h2>{c.name} · {DAYS[editingCell.day]} · {editingCell.period + 1}ª aula</h2></div><button className="icon-button" onClick={() => setEditingCell(null)}><X size={19} /></button></div><p>Escolha uma disciplina. A disponibilidade e os conflitos serão verificados automaticamente.</p><div className="modal-options">{c.subjects.map((s) => { const [bg, border, color] = COLORS[s.color]; return <button className={s.id === currentId ? 'selected' : ''} style={{ background: bg, borderColor: border, color }} key={s.id} onClick={() => manualChange(s.id)}><span><strong>{s.name}</strong><small>{s.teacher}</small></span>{s.id === currentId && <Check size={18} />}</button> })}</div></div></div>
      })()}
    </div>
  )
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return <label className="field"><span>{label}</span>{children}</label>
}

function CardHeading({ icon, title, subtitle }: { icon: React.ReactNode; title: string; subtitle: string }) {
  return <div className="card-heading"><span>{icon}</span><div><h2>{title}</h2><p>{subtitle}</p></div></div>
}

function Empty({ text }: { text: string }) {
  return <div className="empty"><span><BookOpen size={22} /></span><p>{text}</p></div>
}

function BlockedGrid({ value, onChange, lessons }: { value: string[]; onChange: (items: string[]) => void; lessons: number }) {
  return <div className="blocked-grid" style={{ gridTemplateColumns: `36px repeat(${DAYS.length}, minmax(30px, 1fr))` }}><div /><>{DAYS_SHORT.map((d) => <strong key={d}>{d}</strong>)}</>{Array.from({ length: lessons }, (_, p) => <div className="blocked-row" key={p}><small>{p + 1}ª</small>{DAYS_SHORT.map((_, d) => { const key = `${d}-${p}`; return <button type="button" aria-label={`${value.includes(key) ? 'Liberar' : 'Bloquear'} ${DAYS[d]}, ${p + 1}ª aula`} className={value.includes(key) ? 'blocked' : ''} onClick={() => onChange(value.includes(key) ? value.filter((x) => x !== key) : [...value, key])} key={key}>{value.includes(key) && <X size={12} />}</button> })}</div>)}</div>
}

function Schedule({ item, grid, onCell }: { item: SchoolClass; grid?: string[][]; onCell: (day: number, period: number) => void }) {
  return <section className="schedule-card"><div className="schedule-title"><div><span className="class-icon dark"><GraduationCap size={19} /></span><div><h2>{item.name}</h2><p>{item.shift} · {classCapacity(item)} aulas · {item.lessonDuration} min · {item.shift === 'Integral' ? `manhã ${item.startTime} / tarde ${item.afternoonStartTime}` : `início ${item.startTime}`}</p></div></div><span>Grade completa <Check size={15} /></span></div><div className="schedule-scroll"><div className="schedule-grid" style={{ gridTemplateColumns: `120px repeat(${DAYS.length}, minmax(155px, 1fr))` }}><div className="schedule-head time">Período</div>{DAYS.map((d, day) => <div className="schedule-head" key={d}>{d}<small>{item.lessonsByDay[day]} aulas</small></div>)}{Array.from({ length: maxLessons(item) }, (_, period) => <div className="schedule-row" key={period}><div className="period"><strong>{period + 1}ª aula</strong><small>{item.shift === 'Integral' ? 'Manhã / tarde' : item.shift}</small></div>{DAYS.map((_, day) => { if (period >= item.lessonsByDay[day]) return <div className="lesson unavailable-cell" key={day}>Sem aula</div>; const s = item.subjects.find((x) => x.id === grid?.[day]?.[period]); if (!s) return <div className="lesson empty-cell" key={day}>Não alocada</div>; const [bg, border, color] = COLORS[s.color]; return <button className="lesson" style={{ background: bg, borderColor: border, color }} onClick={() => onCell(day, period)} key={day}><span className="lesson-time">{periodLabel(item, day, period)}</span><strong>{s.name}</strong><small>{s.teacher}</small><Edit3 size={13} /></button> })}</div>)}</div></div></section>
}

function safeName(value: string) {
  return value.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '')
}

function downloadBlob(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = name
  a.click()
  URL.revokeObjectURL(url)
}

export default App
