interface MessagesProps {
  items: Array<{ type: 'success' | 'error'; text: string }>
}

export function Messages({ items }: MessagesProps) {
  if (items.length === 0) {
    return null
  }

  return (
    <section className="messages">
      {items.map((item, index) => (
        <article key={`${item.type}-${index}`} className={`message ${item.type}`}>
          {item.text}
        </article>
      ))}
    </section>
  )
}
