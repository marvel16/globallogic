#ifndef THREAD_POOL_H
#define THREAD_POOL_H

#include <vector>
#include <queue>
#include <memory>
#include <thread>
#include <mutex>
#include <condition_variable>
#include <future>
#include <functional>
#include <stdexcept>
#include <fstream>

class ThreadPool {
public:
	ThreadPool(size_t);
	template<class F, class... Args>
	auto enqueue(F&& f, Args&&... args)
		->std::future<typename std::result_of<F(Args...)>::type>;
	~ThreadPool();
private:
	// need to keep track of threads so we can join them
	std::vector< std::thread > workers;
	// the task queue
	std::queue< std::function<void()> > tasks;

	// synchronization
	std::mutex queue_mutex;
	std::condition_variable condition;
	bool stop;
};



// the constructor just launches some amount of workers
inline ThreadPool::ThreadPool(size_t threads)
: stop(false)
{
	for (size_t i = 0; i<threads; ++i)
		workers.emplace_back(
		[this]
	{
		for (;;)
		{
			std::unique_lock<std::mutex> lock(this->queue_mutex);
			while (!this->stop && this->tasks.empty())
				this->condition.wait(lock);
			if (this->stop && this->tasks.empty())
				return;
			std::function<void()> task(this->tasks.front());
			this->tasks.pop();
			lock.unlock();
			task();
		}
	}
	);
}


// add new work item to the pool
template<class F, class... Args>
auto ThreadPool::enqueue(F&& f, Args&&... args)
-> std::future<typename std::result_of<F(Args...)>::type>
{
	typedef typename std::result_of<F(Args...)>::type return_type;

	// don't allow enqueueing after stopping the pool
	if (stop)
		throw std::runtime_error("enqueue on stopped ThreadPool");

	auto task = std::make_shared< std::packaged_task<return_type()> >(
		std::bind(std::forward<F>(f), std::forward<Args>(args)...)
		);

	std::future<return_type> res = task->get_future();
	{
		std::unique_lock<std::mutex> lock(queue_mutex);
		tasks.push([task](){ (*task)(); });
	}
	condition.notify_one();
	return res;
}

// the destructor joins all threads
inline ThreadPool::~ThreadPool()
{
	{
		std::unique_lock<std::mutex> lock(queue_mutex);
		stop = true;
	}
	condition.notify_all();
	for (size_t i = 0; i < workers.size(); ++i)
	{
		if (workers[i].joinable()) workers[i].join();
	}
}


class CheckSum
{
	std::mutex m;
public:
	std::pair<unsigned int, std::wstring>  task(std::pair<unsigned int, std::wstring> &Fpair)
	{
		std::lock_guard<std::mutex> lock(m);
		std::fstream file(Fpair.second, std::ios_base::in | std::ios_base::binary);
		unsigned int sum = 0;
		char buff;

		if (!file.is_open())
		{
			//cerr << "\nFile wasnt opened...\n";
			throw std::runtime_error("(CheckSum::hashSum : File wasnt opened...");
		}
		
		while (file.read(&buff, 1))
			sum += (unsigned int)(buff);
		return Fpair;
	}

	
};

#endif



